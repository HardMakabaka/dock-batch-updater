"""Batch processor for handling multiple DOCX documents.

This module provides functionality for processing multiple documents
simultaneously with progress tracking and error handling.
"""

import os
import threading
from typing import List, Dict, Any, Optional, Callable, Tuple
from pathlib import Path
from queue import Queue, Empty
from concurrent.futures import ThreadPoolExecutor, as_completed

from .docx_processor import DocxProcessor


class ProcessingResult:
    """Result of processing a single document."""

    def __init__(self, file_path: str, success: bool, message: str = "",
                 replacements: int = 0, backup_path: str = ""):
        """Initialize processing result.

        Args:
            file_path: Path to the processed file
            success: Whether processing was successful
            message: Status message
            replacements: Number of replacements made
            backup_path: Path to backup file
        """
        self.file_path = file_path
        self.success = success
        self.message = message
        self.replacements = replacements
        self.backup_path = backup_path

    def to_dict(self) -> Dict[str, Any]:
        """Convert result to dictionary.

        Returns:
            Dictionary representation of the result
        """
        return {
            'file_path': self.file_path,
            'success': self.success,
            'message': self.message,
            'replacements': self.replacements,
            'backup_path': self.backup_path
        }


class BatchProcessor:
    """Process multiple DOCX documents with threading support.

    This class manages batch processing of multiple DOCX documents,
    providing progress tracking, error handling, and result aggregation.
    """

    def __init__(self, max_workers: int = 4):
        """Initialize batch processor.

        Args:
            max_workers: Maximum number of worker threads
        """
        self.max_workers = max_workers
        self.results: List[ProcessingResult] = []
        self._stop_event = threading.Event()
        self._progress_queue: Queue = Queue()

    def process_documents(
        self,
        file_paths: List[str],
        replacements: List[Tuple[str, str]],
        create_backup: bool = True,
        backup_dir: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        result_callback: Optional[Callable[[ProcessingResult], None]] = None
    ) -> List[ProcessingResult]:
        """Process multiple documents with text replacements.

        Args:
            file_paths: List of file paths to process
            replacements: List of (search_text, replace_text) tuples
            create_backup: Whether to create backup files
            backup_dir: Optional directory for backups
            progress_callback: Optional callback for progress updates
            result_callback: Optional callback for individual document results

        Returns:
            List of ProcessingResult objects
        """
        self.results = []
        self._stop_event.clear()

        # Validate all files first
        valid_files = self._validate_files(file_paths)
        invalid_files = set(file_paths) - set(valid_files)

        # Report invalid files
        for file_path in invalid_files:
            result = ProcessingResult(
                file_path=file_path,
                success=False,
                message="Invalid or unreadable DOCX file"
            )
            self.results.append(result)
            if result_callback:
                result_callback(result)

        total_files = len(valid_files)
        processed_count = 0

        # Process valid files
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks
            future_to_file = {
                executor.submit(
                    self._process_single_file,
                    file_path,
                    replacements,
                    create_backup,
                    backup_dir
                ): file_path
                for file_path in valid_files
            }

            # Process results as they complete
            for future in as_completed(future_to_file):
                if self._stop_event.is_set():
                    break

                file_path = future_to_file[future]
                try:
                    result = future.result()
                    self.results.append(result)

                    processed_count += 1

                    if result_callback:
                        result_callback(result)

                    if progress_callback:
                        progress_callback(processed_count, total_files)

                except Exception as e:
                    error_result = ProcessingResult(
                        file_path=file_path,
                        success=False,
                        message=f"Processing error: {str(e)}"
                    )
                    self.results.append(error_result)

                    if result_callback:
                        result_callback(error_result)

                    processed_count += 1

                    if progress_callback:
                        progress_callback(processed_count, total_files)

        return self.results

    def _process_single_file(
        self,
        file_path: str,
        replacements: List[Tuple[str, str]],
        create_backup: bool,
        backup_dir: Optional[str]
    ) -> ProcessingResult:
        """Process a single document.

        Args:
            file_path: Path to the document
            replacements: List of (search_text, replace_text) tuples
            create_backup: Whether to create backup
            backup_dir: Optional backup directory

        Returns:
            ProcessingResult object
        """
        try:
            # Check for stop event
            if self._stop_event.is_set():
                return ProcessingResult(
                    file_path=file_path,
                    success=False,
                    message="Processing cancelled"
                )

            processor = DocxProcessor(file_path)

            # Load document
            if not processor.load():
                return ProcessingResult(
                    file_path=file_path,
                    success=False,
                    message="Failed to load document"
                )

            # Validate document
            is_valid, errors = processor.validate_document()
            if not is_valid:
                return ProcessingResult(
                    file_path=file_path,
                    success=False,
                    message=f"Invalid document: {', '.join(errors)}"
                )

            # Create backup if requested
            backup_path = ""
            if create_backup:
                backup_path = processor.create_backup(backup_dir)

            # Perform replacements
            total_replacements = 0
            for search_text, replace_text in replacements:
                count = processor.replace_text(search_text, replace_text)
                total_replacements += count

            # Save document
            if not processor.save():
                return ProcessingResult(
                    file_path=file_path,
                    success=False,
                    message="Failed to save document",
                    backup_path=backup_path
                )

            return ProcessingResult(
                file_path=file_path,
                success=True,
                message="Successfully processed",
                replacements=total_replacements,
                backup_path=backup_path
            )

        except Exception as e:
            return ProcessingResult(
                file_path=file_path,
                success=False,
                message=f"Error: {str(e)}"
            )

    def _validate_files(self, file_paths: List[str]) -> List[str]:
        """Validate that all files are valid DOCX files.

        Args:
            file_paths: List of file paths to validate

        Returns:
            List of valid file paths
        """
        valid_files = []

        for file_path in file_paths:
            # Check if file exists
            if not os.path.exists(file_path):
                continue

            # Check if file is readable
            if not os.access(file_path, os.R_OK):
                continue

            # Check if it's a DOCX file
            if DocxProcessor.is_docx_file(file_path):
                valid_files.append(file_path)

        return valid_files

    def stop(self) -> None:
        """Stop all processing."""
        self._stop_event.set()

    def get_summary(self) -> Dict[str, Any]:
        """Get a summary of batch processing results.

        Returns:
            Dictionary containing summary statistics
        """
        total = len(self.results)
        successful = sum(1 for r in self.results if r.success)
        failed = total - successful
        total_replacements = sum(r.replacements for r in self.results)

        return {
            'total_files': total,
            'successful': successful,
            'failed': failed,
            'total_replacements': total_replacements,
            'success_rate': successful / total if total > 0 else 0
        }

    def get_failed_results(self) -> List[ProcessingResult]:
        """Get all failed processing results.

        Returns:
            List of failed ProcessingResult objects
        """
        return [r for r in self.results if not r.success]

    def get_successful_results(self) -> List[ProcessingResult]:
        """Get all successful processing results.

        Returns:
            List of successful ProcessingResult objects
        """
        return [r for r in self.results if r.success]

    @staticmethod
    def get_files_from_directory(directory: str, recursive: bool = True) -> List[str]:
        """Get all DOCX files from a directory.

        Args:
            directory: Path to the directory
            recursive: Whether to search recursively

        Returns:
            List of DOCX file paths
        """
        docx_files = []

        if recursive:
            for root, _, files in os.walk(directory):
                for file in files:
                    if file.lower().endswith('.docx'):
                        docx_files.append(os.path.join(root, file))
        else:
            for file in os.listdir(directory):
                if file.lower().endswith('.docx'):
                    docx_files.append(os.path.join(directory, file))

        return sorted(docx_files)

    def clear_results(self) -> None:
        """Clear all processing results."""
        self.results = []
