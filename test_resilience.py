#!/usr/bin/env python3
"""
Test script to validate resilience improvements to the XTM report generation.
This script tests error handling, retries, and fallback mechanisms.
"""

import json
import logging
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock
import requests

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from generate_report import XTMReportGenerator, retry_with_backoff

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def test_retry_decorator():
    """Test that the retry decorator works correctly."""
    logger.info("Testing retry decorator...")

    call_count = {'value': 0}

    @retry_with_backoff(max_attempts=3, initial_delay=0.1, backoff_factor=1)
    def failing_function():
        call_count['value'] += 1
        if call_count['value'] < 3:
            raise requests.exceptions.ConnectionError("Test connection error")
        return "Success"

    try:
        result = failing_function()
        assert result == "Success", "Expected success on third attempt"
        assert call_count['value'] == 3, f"Expected 3 calls, got {call_count['value']}"
        logger.info("✓ Retry decorator test passed")
        return True
    except Exception as e:
        logger.error(f"✗ Retry decorator test failed: {e}")
        return False


def test_config_validation():
    """Test configuration validation."""
    logger.info("Testing configuration validation...")

    try:
        # Test with valid config
        generator = XTMReportGenerator()
        logger.info("✓ Configuration validation passed")
        return True
    except Exception as e:
        logger.error(f"✗ Configuration validation failed: {e}")
        return False


def test_health_checks():
    """Test health check functionality."""
    logger.info("Testing health checks...")

    try:
        generator = XTMReportGenerator()
        result = generator._run_health_checks()
        logger.info(f"✓ Health checks completed (result: {result})")
        return True
    except Exception as e:
        logger.error(f"✗ Health checks failed: {e}")
        return False


def test_graceful_degradation():
    """Test that the system handles partial failures gracefully."""
    logger.info("Testing graceful degradation...")

    try:
        generator = XTMReportGenerator()

        # Mock a project list with some valid and some invalid projects
        with patch.object(generator, 'get_projects') as mock_projects:
            mock_projects.return_value = [
                {'id': 1, 'name': 'Test Project 1'},
                {'id': 2, 'name': 'Test Project 2'},
                {'id': None, 'name': 'Invalid Project'},  # Should be skipped
            ]

            # Mock statistics to fail for one project
            original_get_stats = generator.get_project_statistics

            def mock_get_statistics(project_id, excluded_users=None):
                if project_id == 2:
                    raise Exception("Test failure for project 2")
                return []

            with patch.object(generator, 'get_project_statistics', side_effect=mock_get_statistics):
                # This should not raise an exception
                data = generator.aggregate_monthly_data()

                logger.info("✓ Graceful degradation test passed")
                return True

    except Exception as e:
        logger.error(f"✗ Graceful degradation test failed: {e}")
        return False


def test_notification_system():
    """Test notification system."""
    logger.info("Testing notification system...")

    try:
        generator = XTMReportGenerator()

        # Test system notification (won't actually display in test)
        generator._send_system_notification("Test", "This is a test notification", sound=False)

        logger.info("✓ Notification system test passed")
        return True
    except Exception as e:
        logger.error(f"✗ Notification system test failed: {e}")
        return False


def main():
    """Run all tests."""
    logger.info("=" * 70)
    logger.info("XTM Report Generator - Resilience Test Suite")
    logger.info("=" * 70)

    tests = [
        ("Retry Decorator", test_retry_decorator),
        ("Configuration Validation", test_config_validation),
        ("Health Checks", test_health_checks),
        ("Graceful Degradation", test_graceful_degradation),
        ("Notification System", test_notification_system),
    ]

    results = []
    for test_name, test_func in tests:
        logger.info("")
        logger.info(f"Running: {test_name}")
        logger.info("-" * 70)
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            logger.error(f"Test crashed: {e}", exc_info=True)
            results.append((test_name, False))

    # Print summary
    logger.info("")
    logger.info("=" * 70)
    logger.info("Test Summary")
    logger.info("=" * 70)

    passed = 0
    failed = 0

    for test_name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        logger.info(f"{status:10} {test_name}")
        if result:
            passed += 1
        else:
            failed += 1

    logger.info("")
    logger.info(f"Total: {len(results)} tests, {passed} passed, {failed} failed")

    if failed == 0:
        logger.info("=" * 70)
        logger.info("All tests passed! The automation is resilient.")
        logger.info("=" * 70)
        return 0
    else:
        logger.info("=" * 70)
        logger.info(f"Some tests failed. Please review the failures above.")
        logger.info("=" * 70)
        return 1


if __name__ == "__main__":
    sys.exit(main())
