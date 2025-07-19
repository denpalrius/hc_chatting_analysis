#!/usr/bin/env python3
"""
Test script for the Healthcare Schedule Balancer
"""

import sys
import os
from data_validator import DataValidator
from schedule_balancer import ScheduleBalancer
from excel_formatter import ExcelFormatter

def test_with_sample_file():
    """Test the application with the sample July file"""
    print("🧪 Testing Healthcare Schedule Balancer")
    print("=" * 50)
    
    # Path to sample file
    sample_file = "../data/july_summary.xlsx"
    
    if not os.path.exists(sample_file):
        print(f"❌ Sample file not found: {sample_file}")
        return False
    
    try:
        print("📁 Loading sample file...")
        validator = DataValidator()
        
        # Test day block extraction
        day_blocks, workbook = validator.extract_day_blocks_with_formatting(sample_file)
        print(f"✅ Extracted {len(day_blocks)} day blocks")
        
        # Print first few days for verification
        for i, day in enumerate(day_blocks[:3]):
            print(f"  Day {i+1}: {day['date']} - {len(day['providers'])} providers")
        
        if not day_blocks:
            print("❌ No day blocks found")
            return False
        
        # Test validation
        print("🔍 Validating day blocks...")
        issues = validator.validate_day_blocks(day_blocks)
        print(f"✅ Found {len(issues)} validation issues")
        
        # Test balancing
        print("⚖️ Testing schedule balancing...")
        balancer = ScheduleBalancer()
        processed_workbook, summary = balancer.balance_schedule(day_blocks, workbook)
        
        print("📊 Summary:")
        for key, value in summary.items():
            print(f"  {key}: {value}")
        
        # Test Excel formatting
        print("📋 Testing Excel formatting...")
        formatter = ExcelFormatter()
        changes_log = balancer.get_changes_log()
        print(f"✅ Generated {len(changes_log)} change log entries")
        
        if processed_workbook:
            excel_data = formatter.create_formatted_excel(processed_workbook, changes_log)
            if excel_data:
                print(f"✅ Generated Excel file ({len(excel_data)} bytes)")
            else:
                print("❌ Failed to generate Excel file")
                return False
        
        print("🎉 All tests passed!")
        return True
        
    except Exception as e:
        print(f"❌ Test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Main test function"""
    print("Healthcare Schedule Balancer - Test Suite")
    print("=" * 60)
    
    success = test_with_sample_file()
    
    if success:
        print("\n✅ All tests completed successfully!")
        print("\n🚀 Ready to run: streamlit run app.py")
    else:
        print("\n❌ Tests failed. Please check the errors above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
