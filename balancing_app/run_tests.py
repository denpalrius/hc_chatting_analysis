#!/usr/bin/env python3
"""
Test runner for Schedule Balancer application.

This script runs both unit tests and integration tests and provides
a comprehensive test report.
"""

import sys
import unittest
import time
from io import StringIO


def run_test_suite():
    """Run all test suites and return results"""
    
    print("=" * 60)
    print("Schedule Balancer Test Suite")
    print("=" * 60)
    
    # Import test modules
    try:
        import test_schedule_balancer
        import test_integration
    except ImportError as e:
        print(f"Error importing test modules: {e}")
        print("Make sure all required modules are available.")
        return False
    
    # Create test suites
    unit_suite = unittest.TestLoader().loadTestsFromModule(test_schedule_balancer)
    integration_suite = unittest.TestLoader().loadTestsFromModule(test_integration)
    
    # Combine suites
    combined_suite = unittest.TestSuite([unit_suite, integration_suite])
    
    # Run tests with detailed output
    stream = StringIO()
    runner = unittest.TextTestRunner(
        stream=stream, 
        verbosity=2,
        buffer=True
    )
    
    print("\nRunning Unit Tests...")
    print("-" * 40)
    
    start_time = time.time()
    unit_result = unittest.TextTestRunner(verbosity=2, buffer=True).run(unit_suite)
    unit_time = time.time() - start_time
    
    print(f"\nUnit Tests completed in {unit_time:.2f} seconds")
    print(f"Tests run: {unit_result.testsRun}")
    print(f"Failures: {len(unit_result.failures)}")
    print(f"Errors: {len(unit_result.errors)}")
    
    if unit_result.failures:
        print("\nUnit Test Failures:")
        for test, traceback in unit_result.failures:
            print(f"- {test}: {traceback}")
    
    if unit_result.errors:
        print("\nUnit Test Errors:")
        for test, traceback in unit_result.errors:
            print(f"- {test}: {traceback}")
    
    print("\n" + "=" * 60)
    print("Running Integration Tests...")
    print("-" * 40)
    
    start_time = time.time()
    integration_result = unittest.TextTestRunner(verbosity=2, buffer=True).run(integration_suite)
    integration_time = time.time() - start_time
    
    print(f"\nIntegration Tests completed in {integration_time:.2f} seconds")
    print(f"Tests run: {integration_result.testsRun}")
    print(f"Failures: {len(integration_result.failures)}")
    print(f"Errors: {len(integration_result.errors)}")
    
    if integration_result.failures:
        print("\nIntegration Test Failures:")
        for test, traceback in integration_result.failures:
            print(f"- {test}: {traceback}")
    
    if integration_result.errors:
        print("\nIntegration Test Errors:")
        for test, traceback in integration_result.errors:
            print(f"- {test}: {traceback}")
    
    # Summary
    total_tests = unit_result.testsRun + integration_result.testsRun
    total_failures = len(unit_result.failures) + len(integration_result.failures)
    total_errors = len(unit_result.errors) + len(integration_result.errors)
    total_time = unit_time + integration_time
    
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    print(f"Total tests run: {total_tests}")
    print(f"Total failures: {total_failures}")
    print(f"Total errors: {total_errors}")
    print(f"Total time: {total_time:.2f} seconds")
    
    success = total_failures == 0 and total_errors == 0
    
    if success:
        print("\n✅ ALL TESTS PASSED!")
    else:
        print("\n❌ SOME TESTS FAILED!")
        print(f"Success rate: {((total_tests - total_failures - total_errors) / total_tests * 100):.1f}%")
    
    return success


def run_specific_test_class(test_class_name):
    """Run tests from a specific test class"""
    
    print(f"Running tests from class: {test_class_name}")
    print("-" * 40)
    
    # Import modules
    try:
        import test_schedule_balancer
        import test_integration
    except ImportError as e:
        print(f"Error importing test modules: {e}")
        return False
    
    # Find the test class
    test_class = None
    if hasattr(test_schedule_balancer, test_class_name):
        test_class = getattr(test_schedule_balancer, test_class_name)
    elif hasattr(test_integration, test_class_name):
        test_class = getattr(test_integration, test_class_name)
    
    if not test_class:
        print(f"Test class '{test_class_name}' not found")
        available_classes = []
        for module in [test_schedule_balancer, test_integration]:
            for name in dir(module):
                obj = getattr(module, name)
                if isinstance(obj, type) and issubclass(obj, unittest.TestCase):
                    available_classes.append(name)
        print(f"Available test classes: {', '.join(available_classes)}")
        return False
    
    # Run the specific test class
    suite = unittest.TestLoader().loadTestsFromTestCase(test_class)
    result = unittest.TextTestRunner(verbosity=2, buffer=True).run(suite)
    
    success = len(result.failures) == 0 and len(result.errors) == 0
    print(f"\nResult: {result.testsRun} tests, {len(result.failures)} failures, {len(result.errors)} errors")
    
    return success


def run_specific_test_method(test_class_name, test_method_name):
    """Run a specific test method"""
    
    print(f"Running test: {test_class_name}.{test_method_name}")
    print("-" * 40)
    
    try:
        import test_schedule_balancer
        import test_integration
    except ImportError as e:
        print(f"Error importing test modules: {e}")
        return False
    
    # Find the test class
    test_class = None
    if hasattr(test_schedule_balancer, test_class_name):
        test_class = getattr(test_schedule_balancer, test_class_name)
    elif hasattr(test_integration, test_class_name):
        test_class = getattr(test_integration, test_class_name)
    
    if not test_class:
        print(f"Test class '{test_class_name}' not found")
        return False
    
    # Run the specific test method
    suite = unittest.TestSuite()
    suite.addTest(test_class(test_method_name))
    result = unittest.TextTestRunner(verbosity=2, buffer=True).run(suite)
    
    success = len(result.failures) == 0 and len(result.errors) == 0
    print(f"\nResult: {result.testsRun} tests, {len(result.failures)} failures, {len(result.errors)} errors")
    
    return success


def main():
    """Main function to handle command line arguments"""
    
    if len(sys.argv) == 1:
        # Run all tests
        success = run_test_suite()
        sys.exit(0 if success else 1)
    
    elif len(sys.argv) == 2:
        # Run specific test class
        test_class = sys.argv[1]
        success = run_specific_test_class(test_class)
        sys.exit(0 if success else 1)
    
    elif len(sys.argv) == 3:
        # Run specific test method
        test_class = sys.argv[1]
        test_method = sys.argv[2]
        success = run_specific_test_method(test_class, test_method)
        sys.exit(0 if success else 1)
    
    else:
        print("Usage:")
        print("  python run_tests.py                    # Run all tests")
        print("  python run_tests.py TestClassName      # Run tests from specific class")
        print("  python run_tests.py TestClass method   # Run specific test method")
        print("")
        print("Available test classes:")
        print("  - TestScheduleBalancer (unit tests)")
        print("  - TestIntegration (integration tests)")
        sys.exit(1)


if __name__ == '__main__':
    main()
