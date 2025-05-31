#!/bin/bash
# Clean Python cache files and debug logs

echo "Cleaning Python cache files..."
find . -type d -name "__pycache__" -exec rm -rf {} +
find . -type f -name "*.pyc" -delete
find . -type f -name "*.pyo" -delete
find . -type f -name "*.pyd" -delete
find . -type f -name ".DS_Store" -delete

echo "Cleaning debug logs and output files..."
rm -f razorpay_*.json
rm -f debug_output.json
rm -f razorpay_sync.log

echo "Done! Cache files and debug logs have been removed." 