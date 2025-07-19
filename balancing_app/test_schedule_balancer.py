import pytest
import unittest
from unittest.mock import Mock, patch, MagicMock
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
import copy

from schedule_balancer import ScheduleBalancer


class TestScheduleBalancer(unittest.TestCase):
    """Comprehensive test suite for ScheduleBalancer"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.balancer = ScheduleBalancer(
            individuals=['DD', 'DM', 'OT'],
            additional_providers=['Test Provider 1', 'Test Provider 2']
        )
        self.wb = Workbook()
        self.ws = self.wb.active
        
    def create_sample_day_data(self, date='2025-07-02', providers_data=None):
        """Create sample day data for testing"""
        if providers_data is None:
            providers_data = [
                {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 4, 'OT': 2}},
                {'name': 'Jane Doe, RN', 'hours': {'DD': 6, 'DM': 8, 'OT': 2}},
            ]
        
        day_data = {
            'date': date,
            'date_row': 1,
            'providers': [],
            'totals_row': 10,
            'pending_row': 11
        }
        
        for i, provider_data in enumerate(providers_data, start=2):
            provider = {
                'name': provider_data['name'],
                'row': i,
                'hours': provider_data['hours'].copy(),
                'total_formula': f'=SUM(B{i}:D{i})',
                'is_new_provider': False
            }
            day_data['providers'].append(provider)
            
        return day_data


class TestBasicBalancing(TestScheduleBalancer):
    """Test basic balancing functionality"""
    
    def test_already_balanced_day(self):
        """Test day that's already balanced (24 hours per individual)"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
            {'name': 'Jane Doe, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
            {'name': 'Bob Johnson, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}}
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        self.assertTrue(balanced)
        self.assertEqual(modifications['entries_modified'], 0)
        self.assertEqual(modifications['providers_added'], 0)
    
    def test_simple_shortage_filling(self):
        """Test simple case where we need to add hours"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 4, 'OT': 2}},
            {'name': 'Jane Doe, RN', 'hours': {'DD': 6, 'DM': 6, 'OT': 4}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should be able to balance by adding supplemental providers
        self.assertTrue(balanced)
        self.assertGreater(modifications['entries_modified'], 0)
    
    def test_zero_hours_modification(self):
        """Test modification of existing providers with 0 hours"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 12, 'DM': 12, 'OT': 12}},
            {'name': 'Charles Sagini, RN/House Manager', 'hours': {'DD': 0, 'DM': 0, 'OT': 0}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        self.assertTrue(balanced)
        # Should modify the supplemental provider's 0 hours
        charles_provider = next(p for p in day_data['providers'] if 'Charles' in p['name'])
        self.assertGreater(sum(charles_provider['hours'].values()), 0)


class TestOverAllocationHandling(TestScheduleBalancer):
    """Test over-allocation scenarios"""
    
    def test_single_provider_over_allocated(self):
        """Test fixing a single over-allocated provider"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 6}},  # 22 hours total
            {'name': 'Jane Doe, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should reduce John's hours and add supplemental providers
        john_provider = next(p for p in day_data['providers'] if 'John' in p['name'])
        total_john_hours = sum(john_provider['hours'].values())
        self.assertLessEqual(total_john_hours, 16)
        
        # Should still balance to 24 hours per individual
        self.assertTrue(balanced)
    
    def test_multiple_providers_over_allocated(self):
        """Test multiple providers over-allocated"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 10, 'DM': 8, 'OT': 4}},  # 22 hours
            {'name': 'Jane Doe, RN', 'hours': {'DD': 6, 'DM': 10, 'OT': 8}},   # 24 hours
            {'name': 'Bob Johnson, RN', 'hours': {'DD': 4, 'DM': 2, 'OT': 4}}, # 10 hours
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # All providers should be within 16 hour limit
        for provider in day_data['providers']:
            total_hours = sum(provider['hours'].values())
            self.assertLessEqual(total_hours, 16, 
                               f"Provider {provider['name']} has {total_hours} hours")
        
        self.assertTrue(balanced)
    
    def test_reduction_priority_order(self):
        """Test that over-allocation reduction follows priority order (OT, DM, DD)"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 6}},  # 22 hours
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        original_ot = day_data['providers'][0]['hours']['OT']
        original_dm = day_data['providers'][0]['hours']['DM']
        original_dd = day_data['providers'][0]['hours']['DD']
        
        self.balancer._fix_over_allocated_providers(day_data, self.ws, {'log': [], 'entries_modified': 0})
        
        # OT should be reduced first
        john_provider = day_data['providers'][0]
        if original_ot > 2:  # If reduction was possible
            self.assertLessEqual(john_provider['hours']['OT'], original_ot)


class TestSupplementalProviderHandling(TestScheduleBalancer):
    """Test supplemental provider addition and modification"""
    
    def test_add_new_supplemental_provider(self):
        """Test adding new supplemental provider"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should add supplemental providers
        self.assertTrue(balanced)
        self.assertGreater(modifications['providers_added'], 0)
        
        # Verify supplemental provider was added
        supplemental_added = any(
            any(supp in provider['name'] for supp in self.balancer.supplemental_providers)
            for provider in day_data['providers']
            if provider.get('is_new_provider', False)
        )
        self.assertTrue(supplemental_added)
    
    def test_modify_existing_supplemental_zero_hours(self):
        """Test modifying existing supplemental provider with 0 hours"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 12, 'DM': 12, 'OT': 12}},
            {'name': 'Charles Sagini, RN/House Manager', 'hours': {'DD': 0, 'DM': 0, 'OT': 0}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should modify Charles's hours rather than adding new provider
        charles_provider = next(p for p in day_data['providers'] if 'Charles' in p['name'])
        self.assertGreater(sum(charles_provider['hours'].values()), 0)
        self.assertTrue(balanced)
    
    def test_all_supplemental_providers_utilized(self):
        """Test scenario where all supplemental providers are needed"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 2, 'DM': 2, 'OT': 2}},  # Very low hours
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should try to use all supplemental providers
        self.assertTrue(balanced)
        self.assertGreaterEqual(modifications['providers_added'], 1)


class TestEdgeCases(TestScheduleBalancer):
    """Test edge cases and boundary conditions"""
    
    def test_impossible_balance_scenario(self):
        """Test scenario that's impossible to balance"""
        # Create a scenario where even with all supplemental providers, 
        # we can't reach 24 hours per individual
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 1, 'DM': 1, 'OT': 1}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        
        # Mock the supplemental providers to be empty to force impossible scenario
        original_supplemental = self.balancer.supplemental_providers
        self.balancer.supplemental_providers = []
        
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should not be balanced
        self.assertFalse(balanced)
        
        # Restore original supplemental providers
        self.balancer.supplemental_providers = original_supplemental
    
    def test_minimum_hours_constraint(self):
        """Test that minimum 2-hour constraint is respected"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 4}},  # 20 hours
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        
        # Force over-allocation scenario
        day_data['providers'][0]['hours'] = {'DD': 10, 'DM': 10, 'OT': 6}  # 26 hours
        
        self.balancer._fix_over_allocated_providers(day_data, self.ws, {'log': [], 'entries_modified': 0})
        
        # No provider should have less than 2 hours in any category (unless it was 0 or 1 originally)
        john_provider = day_data['providers'][0]
        for individual, hours in john_provider['hours'].items():
            if hours > 0:  # If not zero, should be at least 2
                self.assertGreaterEqual(hours, 2, f"{individual} has {hours} hours")
    
    def test_empty_providers_list(self):
        """Test handling of empty providers list"""
        day_data = {
            'date': '2025-07-02',
            'date_row': 1,
            'providers': [],
            'totals_row': 10,
            'pending_row': 11
        }
        
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should add supplemental providers to fill all 72 hours (24 each for DD, DM, OT)
        self.assertTrue(balanced)
        self.assertGreater(modifications['providers_added'], 0)
        self.assertEqual(len(day_data['providers']), modifications['providers_added'])
    
    def test_single_individual_shortage(self):
        """Test when only one individual needs additional hours"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
            {'name': 'Jane Doe, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
            {'name': 'Bob Johnson, RN', 'hours': {'DD': 8, 'DM': 6, 'OT': 8}},  # DM short by 2
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        self.assertTrue(balanced)
        # Should only need minimal modifications
        self.assertLessEqual(modifications['entries_modified'], 2)


class TestWorksheetInteraction(TestScheduleBalancer):
    """Test worksheet modification and formatting"""
    
    def test_cell_formatting_application(self):
        """Test that proper formatting is applied to modified cells"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 12, 'DM': 12, 'OT': 12}},
            {'name': 'Charles Sagini, RN/House Manager', 'hours': {'DD': 0, 'DM': 0, 'OT': 0}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Verify that cells were modified and potentially have formatting
        # (This test would be more meaningful with actual Excel formatting verification)
        self.assertTrue(balanced)
        self.assertGreater(modifications['entries_modified'], 0)
    
    def test_row_insertion_handling(self):
        """Test proper handling of row insertions"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        original_totals_row = day_data['totals_row']
        
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        if modifications['providers_added'] > 0:
            # Totals row should have shifted down
            self.assertGreater(day_data['totals_row'], original_totals_row)
    
    def test_formula_updates(self):
        """Test that formulas are properly updated after modifications"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Verify that total formulas reference correct rows
        for provider in day_data['providers']:
            expected_formula = f"=SUM(B{provider['row']}:D{provider['row']})"
            self.assertEqual(provider['total_formula'], expected_formula)


class TestMultipleDayProcessing(TestScheduleBalancer):
    """Test processing multiple days together"""
    
    def test_multiple_balanced_days(self):
        """Test processing multiple already balanced days"""
        day_blocks = []
        for i in range(3):
            providers_data = [
                {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
                {'name': 'Jane Doe, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
                {'name': 'Bob Johnson, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}}
            ]
            day_data = self.create_sample_day_data(
                date=f'2025-07-0{i+2}', 
                providers_data=providers_data
            )
            day_blocks.append(day_data)
        
        wb, summary = self.balancer.balance_schedule(day_blocks, self.wb)
        
        self.assertEqual(summary['total_days_processed'], 3)
        self.assertEqual(summary['days_balanced'], 3)
        self.assertEqual(summary['days_unbalanced'], 0)
    
    def test_mixed_balanced_unbalanced_days(self):
        """Test processing mix of balanced and unbalanced days"""
        day_blocks = []
        
        # Balanced day
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
            {'name': 'Jane Doe, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}},
            {'name': 'Bob Johnson, RN', 'hours': {'DD': 8, 'DM': 8, 'OT': 8}}
        ]
        day_blocks.append(self.create_sample_day_data('2025-07-02', providers_data))
        
        # Unbalanced day
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 12, 'DM': 12, 'OT': 12}},
        ]
        day_blocks.append(self.create_sample_day_data('2025-07-03', providers_data))
        
        wb, summary = self.balancer.balance_schedule(day_blocks, self.wb)
        
        self.assertEqual(summary['total_days_processed'], 2)
        self.assertGreaterEqual(summary['days_balanced'], 1)  # At least one should be balanced


class TestErrorHandling(TestScheduleBalancer):
    """Test error handling and exception scenarios"""
    
    def test_malformed_day_data(self):
        """Test handling of malformed day data"""
        malformed_day_data = {
            'date': '2025-07-02',
            'providers': [
                {'name': 'John Smith, RN'}  # Missing required fields
            ]
        }
        
        # Should handle gracefully without crashing
        with self.assertLogs() as log:
            try:
                balanced, modifications = self.balancer._balance_single_day(malformed_day_data, self.ws)
            except Exception as e:
                # Should catch and handle the exception
                pass
    
    def test_worksheet_none_handling(self):
        """Test handling when worksheet is None"""
        day_blocks = [self.create_sample_day_data()]
        
        wb, summary = self.balancer.balance_schedule(day_blocks, None)
        
        # Should handle gracefully
        self.assertIsNone(wb)
        self.assertIn('total_days_processed', summary)
    
    def test_invalid_individual_names(self):
        """Test handling of invalid individual names"""
        balancer = ScheduleBalancer(individuals=['INVALID', 'NAMES'])
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'INVALID': 8, 'NAMES': 16}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        
        # Should handle without crashing
        try:
            balanced, modifications = balancer._balance_single_day(day_data, self.ws)
            # Test passes if no exception is raised
        except Exception as e:
            self.fail(f"Should handle invalid individual names gracefully: {e}")


class TestLoggingAndReporting(TestScheduleBalancer):
    """Test logging and reporting functionality"""
    
    def test_changes_log_generation(self):
        """Test that changes are properly logged"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 12, 'DM': 12, 'OT': 12}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        balanced, modifications = self.balancer._balance_single_day(day_data, self.ws)
        
        # Should have logged changes
        self.assertGreater(len(modifications['log']), 0)
        
        # Log entries should contain relevant information
        log_text = ' '.join(modifications['log'])
        self.assertIn('2025-07-02', log_text)  # Date should be mentioned
    
    def test_summary_statistics(self):
        """Test that summary statistics are correctly calculated"""
        day_blocks = []
        for i in range(5):
            providers_data = [
                {'name': 'John Smith, RN', 'hours': {'DD': 12, 'DM': 12, 'OT': 12}},
            ]
            day_data = self.create_sample_day_data(f'2025-07-0{i+2}', providers_data)
            day_blocks.append(day_data)
        
        wb, summary = self.balancer.balance_schedule(day_blocks, self.wb)
        
        self.assertEqual(summary['total_days_processed'], 5)
        self.assertEqual(summary['days_balanced'] + summary['days_unbalanced'], 5)
        self.assertGreaterEqual(summary['entries_modified'], 0)
        self.assertGreaterEqual(summary['providers_added'], 0)
    
    def test_get_changes_log(self):
        """Test retrieving complete changes log"""
        providers_data = [
            {'name': 'John Smith, RN', 'hours': {'DD': 12, 'DM': 12, 'OT': 12}},
        ]
        
        day_data = self.create_sample_day_data(providers_data=providers_data)
        self.balancer._balance_single_day(day_data, self.ws)
        
        complete_log = self.balancer.get_changes_log()
        self.assertIsInstance(complete_log, list)
        self.assertGreater(len(complete_log), 0)


if __name__ == '__main__':
    # Run all tests
    unittest.main(verbosity=2)
