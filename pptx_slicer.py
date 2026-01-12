"""
PowerPoint Slicer - Split a PowerPoint file into individual slides
Each slide will be saved as a separate PowerPoint file and as images.
"""

import os
import shutil
from pathlib import Path
from pptx import Presentation
import comtypes.client
from PIL import Image


def export_slides_as_images(input_file, output_dir=None, image_format='png'):
    """
    Export each slide as an image file.
    
    Args:
        input_file (str): Path to the input PowerPoint file
        output_dir (str, optional): Directory to save images. 
                                   Defaults to 'output/images' folder.
        image_format (str): Image format - 'png', 'jpg', or 'jpeg'
    
    Returns:
        list: List of paths to the created image files
    """
    # Set up output directory
    if output_dir is None:
        output_dir = os.path.join("output", "images")
    
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # Convert to absolute path
    input_file = os.path.abspath(input_file)
    output_dir = os.path.abspath(output_dir)
    
    print(f"Exporting slides as images from: {input_file}")
    
    # Initialize PowerPoint
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    
    try:
        # Open the presentation
        presentation = powerpoint.Presentations.Open(input_file, WithWindow=False)
        total_slides = presentation.Slides.Count
        
        created_images = []
        input_name = Path(input_file).stem
        
        # Determine image format constant
        format_map = {
            'png': 'png',
            'jpg': 'jpg',
            'jpeg': 'jpg'
        }
        
        img_ext = format_map.get(image_format.lower(), 'png')
        
        # Export each slide
        for slide_idx in range(1, total_slides + 1):
            slide = presentation.Slides(slide_idx)
            image_path = os.path.join(output_dir, f"{input_name}_slide_{slide_idx}.{img_ext}")
            
            # Export slide as image
            slide.Export(image_path, img_ext.upper())
            created_images.append(image_path)
            print(f"Created image: {image_path}")
        
        # Close presentation
        presentation.Close()
        print(f"\n✓ Successfully created {len(created_images)} image files in '{output_dir}'")
        
        return created_images
        
    finally:
        # Quit PowerPoint
        powerpoint.Quit()


def split_pptx(input_file, output_dir=None):
    """
    Split a PowerPoint file into individual slides.
    
    Args:
        input_file (str): Path to the input PowerPoint file
        output_dir (str, optional): Directory to save output files. 
                                   Defaults to 'output/pptx' folder.
    
    Returns:
        list: List of paths to the created PowerPoint files
    """
    # Validate input file
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    # Set up output directory
    if output_dir is None:
        output_dir = os.path.join("output", "pptx")
    
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # Load the presentation to get slide count
    print(f"Loading presentation: {input_file}")
    prs = Presentation(input_file)
    total_slides = len(prs.slides)
    print(f"Total slides found: {total_slides}")
    
    created_files = []
    input_name = Path(input_file).stem
    
    # Process each slide
    for slide_idx in range(total_slides):
        # Create output filename
        output_file = os.path.join(output_dir, f"{input_name}_slide_{slide_idx + 1}.pptx")
        
        # Create a fresh copy of the presentation for each slide
        # Load the presentation fresh each time
        temp_prs = Presentation(input_file)
        
        # Create list of slide indices to remove (all except the current one)
        slides_to_remove = list(range(total_slides))
        slides_to_remove.remove(slide_idx)
        
        # Remove slides in reverse order to maintain indices
        for idx in reversed(slides_to_remove):
            rId = temp_prs.slides._sldIdLst[idx].rId
            temp_prs.part.drop_rel(rId)
            del temp_prs.slides._sldIdLst[idx]
        
        # Save the presentation with only one slide
        temp_prs.save(output_file)
        created_files.append(output_file)
        print(f"Created: {output_file}")
    
    print(f"\n✓ Successfully created {total_slides} PowerPoint files in '{output_dir}'")
    return created_files


def split_pptx_with_images(input_file, output_pptx_dir=None, output_images_dir=None, image_format='png'):
    """
    Split PowerPoint file into individual slides AND export as images.
    
    Args:
        input_file (str): Path to the input PowerPoint file
        output_pptx_dir (str, optional): Directory to save PPTX files
        output_images_dir (str, optional): Directory to save image files
        image_format (str): Image format - 'png', 'jpg', or 'jpeg'
    
    Returns:
        dict: Dictionary with 'pptx_files' and 'image_files' lists
    """
    print("=" * 60)
    print("PowerPoint Slicer - Generating PPTs and Images")
    print("=" * 60)
    print()
    
    # Generate individual PPTX files
    print("Step 1: Creating individual PPTX files...")
    pptx_files = split_pptx(input_file, output_pptx_dir)
    print()
    
    # Generate images
    print("Step 2: Creating slide images...")
    image_files = export_slides_as_images(input_file, output_images_dir, image_format)
    print()
    
    print("=" * 60)
    print(f"✓ Complete! Generated {len(pptx_files)} PPTX files and {len(image_files)} images")
    print("=" * 60)
    
    return {
        'pptx_files': pptx_files,
        'image_files': image_files,
        'total_slides': len(pptx_files)
    }


def main():
    """Main function to run the PowerPoint slicer."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Split a PowerPoint file into individual slides and images"
    )
    parser.add_argument(
        "input_file",
        help="Path to the input PowerPoint file (.pptx)"
    )
    parser.add_argument(
        "-o", "--output",
        default="output",
        help="Base output directory (default: 'output')"
    )
    parser.add_argument(
        "--with-images",
        action="store_true",
        help="Also generate images for each slide"
    )
    parser.add_argument(
        "--image-format",
        default="png",
        choices=["png", "jpg", "jpeg"],
        help="Image format (default: png)"
    )
    
    args = parser.parse_args()
    
    try:
        if args.with_images:
            # Generate both PPTX and images
            pptx_dir = os.path.join(args.output, "pptx")
            images_dir = os.path.join(args.output, "images")
            split_pptx_with_images(args.input_file, pptx_dir, images_dir, args.image_format)
        else:
            # Generate only PPTX files
            split_pptx(args.input_file, args.output)
    except Exception as e:
        print(f"Error: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
