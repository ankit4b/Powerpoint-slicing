"""
Updated test script for PowerPoint Slicer API v2.0
Demonstrates how to call the API endpoint for PPTX + image generation
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

def test_slice_with_images(file_path, output_pptx_dir=None, output_images_dir=None, image_format="png"):
    """Test the slice endpoint with image generation"""
    print(f"Testing PowerPoint slicing with images for: {file_path}")
    
    payload = {
        "file_path": file_path,
        "generate_images": True,
        "image_format": image_format
    }
    
    if output_pptx_dir:
        payload["output_pptx_dir"] = output_pptx_dir
    
    if output_images_dir:
        payload["output_images_dir"] = output_images_dir
    
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
        
        print(f"\nPPTX Files ({len(result['pptx_files'])}):")
        for file in result['pptx_files'][:3]:  # Show first 3
            print(f"  - {file}")
        if len(result['pptx_files']) > 3:
            print(f"  ... and {len(result['pptx_files']) - 3} more")
        
        if result.get('image_files'):
            print(f"\nImage Files ({len(result['image_files'])}):")
            for file in result['image_files'][:3]:  # Show first 3
                print(f"  - {file}")
            if len(result['image_files']) > 3:
                print(f"  ... and {len(result['image_files']) - 3} more")
    else:
        print(f"Error: {response.json()}")
    
    print()

def test_slice_pptx_only(file_path):
    """Test the slice endpoint for PPTX only (no images)"""
    print(f"Testing PowerPoint slicing (PPTX only) for: {file_path}")
    
    payload = {
        "file_path": file_path,
        "generate_images": False
    }
    
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
        print(f"PPTX Files: {len(result['pptx_files'])}")
    else:
        print(f"Error: {response.json()}")
    
    print()

if __name__ == "__main__":
    print("=" * 60)
    print("PowerPoint Slicer API v2.0 - Test Client")
    print("=" * 60)
    print()
    
    # Test health check
    test_health_check()
    
    # Test slicing with images (PNG format)
    file_path = r"D:\Python\pptx-slicing\demo.pptx"
    print("Test 1: Generate both PPTX and PNG images")
    print("-" * 60)
    test_slice_with_images(file_path)
    
    # Test with JPG format
    print("Test 2: Generate both PPTX and JPG images")
    print("-" * 60)
    test_slice_with_images(file_path, image_format="jpg")
    
    # Test PPTX only
    print("Test 3: Generate PPTX files only (no images)")
    print("-" * 60)
    test_slice_pptx_only(file_path)
