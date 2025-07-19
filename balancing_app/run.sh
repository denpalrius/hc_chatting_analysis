#!/bin/bash

echo "🏥 Healthcare Provider Schedule Balancer"
echo "========================================"

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "❌ Virtual environment not found. Creating one..."
    python3 -m venv venv
    source venv/bin/activate
    pip install --upgrade pip
    pip install -r requirements.txt
else
    echo "✅ Virtual environment found"
    source venv/bin/activate
fi

# Check if all dependencies are installed
echo "📦 Checking dependencies..."
python -c "import streamlit, pandas, openpyxl, xlsxwriter; print('✅ All dependencies installed')"

if [ $? -ne 0 ]; then
    echo "❌ Missing dependencies. Installing..."
    pip install -r requirements.txt
fi

# Run the application
echo ""
echo "🚀 Starting Healthcare Schedule Balancer..."
echo "📱 Your browser should automatically open to the application"
echo "🔗 If not, open: http://localhost:8501"
echo ""
echo "Press Ctrl+C to stop the application"
echo ""

streamlit run app.py
