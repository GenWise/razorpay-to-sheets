#!/usr/bin/env python3
"""
Helper module to load environment variables from .env file
"""

import os
from pathlib import Path

def load_env_vars():
    """Load environment variables from .env file if it exists"""
    try:
        from dotenv import load_dotenv
        
        # Find the .env file in the current directory or parent directories
        env_path = Path('.') / '.env'
        
        if env_path.exists():
            # Load the .env file
            load_dotenv(dotenv_path=env_path)
            print(f"Loaded environment variables from {env_path.absolute()}")
        else:
            print("No .env file found. Using environment variables from system.")
    except ImportError:
        print("Warning: python-dotenv package not installed. Run 'pip install python-dotenv' to use .env files.")
        print("Using environment variables from system.")
        
    # Validate required environment variables
    required_vars = [
        "RAZORPAY_KEY_ID", 
        "RAZORPAY_KEY_SECRET", 
        "GOOGLE_SHEET_ID"
    ]
    
    missing_vars = [var for var in required_vars if not os.environ.get(var)]
    
    if missing_vars:
        print(f"Warning: Missing required environment variables: {', '.join(missing_vars)}")
        print("Please create a .env file or set these variables in your environment.")

if __name__ == "__main__":
    load_env_vars() 