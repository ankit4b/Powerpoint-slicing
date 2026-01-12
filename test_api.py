"""
Test script for PowerPoint Slicer API
Demonstrates how to call the API endpoint
"""

import requests
import json

# API endpoint
BASE_URL = "http://localhost:8000"

def test_health_check():
    """Test the health check endpoint"""
    print("Testing health check...")
    response = requests.get(f"{BASE_URL}/health")
    print(f"Status: {response.status_code}")
    print(f"Response: {response.json()}\n")

def test_slice_pptx(file_path, output_dir=None):
    """Test the slice endpoint"""
    print(f"Testing PowerPoint slicing for: {file_path}")
    
    payload = {
        "file_path": file_path
    }
    
    if output_dir:
        payload["output_dir"] = output_dir
    
    response = requests.post(
        f"{BASE_URL}/slice",
        json=payload,
        headers={"Content-Type": "application/json"}
    )
    
    print(f"Status: {response.status_code}")
    
    if response.status_code == 200:
        result = response.json()
        print(f"Success: {result['success']}")
        print(f"Message: {result['message']}")
        print(f"Total slides: {result['total_slides']}")
        print(f"Output files:")
        for file in result['output_files']:
            print(f"  - {file}")
    else:
        print(f"Error: {response.json()}")
    
    print()

if __name__ == "__main__":
    print("=" * 60)
    print("PowerPoint Slicer API - Test Client")
    print("=" * 60)
    print()
    
    # Test health check
    test_health_check()
    
    # Test slicing with demo.pptx
    # Update this path to your actual file location
    file_path = r"D:\Python\pptx-slicing\demo.pptx"
    test_slice_pptx(file_path)
    
    # Example with custom output directory
    # test_slice_pptx(file_path, output_dir=r"D:\Python\pptx-slicing\custom_output")
