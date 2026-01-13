"""
PowerPoint Slicer - Split a PowerPoint file into individual slides
Each slide will be saved as a separate PowerPoint file and optionally as images.
"""

import os
from pathlib import Path
from pptx import Presentation
from PIL import Image, ImageDraw
import platform
import io


def export_slides_as_images(input_file, output_dir=None, image_format='png'):
    """
    Export each slide as an image file.
    Uses the best available method based on platform and installed packages.
    
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
    
    # Try different methods in order of preference
    
    # Method 1: Windows COM (Windows only)
    if platform.system() == 'Windows':
        try:
            return _export_slides_windows_com(input_file, output_dir, image_format)
        except Exception as e:
            print(f"Windows COM method failed: {e}")
            print("Falling back to pure Python method...")
    
    # Method 2: Using Pillow to extract embedded images from slides (Pure Python - Databricks compatible)
    try:
        return _export_slides_pure_python(input_file, output_dir, image_format)
    except Exception as e:
        print(f"Pure Python method failed: {e}")
        
        # If all methods fail, provide helpful error
        raise RuntimeError(
            "Unable to generate slide images. Options:\n"
            "1. On Windows: Install pywin32 (pip install pywin32)\n"
            "2. On Linux/Databricks: Using pure Python extraction (basic)\n"
            "3. Alternative: Use only PPTX splitting (set generate_images=False)\n"
            "4. For production: Consider pre-processing slides or using a dedicated service"
        )


def _export_slides_pure_python(input_file, output_dir, image_format='png'):
    """
    Export slides using pure Python - extracts images and creates slide previews.
    This is a Databricks-compatible solution that works without external dependencies.
    """
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    
    prs = Presentation(input_file)
    input_name = Path(input_file).stem
    created_images = []
    
    img_ext = image_format.lower()
    if img_ext == 'jpeg':
        img_ext = 'jpg'
    
    print("Using pure Python extraction method (Databricks-compatible)...")
    print("Note: This creates slide backgrounds/images when available.")
    
    for slide_idx, slide in enumerate(prs.slides, start=1):
        image_path = os.path.join(output_dir, f"{input_name}_slide_{slide_idx}.{img_ext}")
        
        # Try to extract the slide background or first image
        slide_image = None
        
        # Look for images in the slide
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # Found an image
                image = shape.image
                image_bytes = image.blob
                slide_image = Image.open(io.BytesIO(image_bytes))
                break
        
        # If no image found, create a placeholder
        if slide_image is None:
            # Create a simple placeholder image with slide number
            slide_image = Image.new('RGB', (1920, 1080), color='white')
            draw = ImageDraw.Draw(slide_image)
            
            # Add slide number
            try:
                text = f"Slide {slide_idx}"
                bbox = draw.textbbox((0, 0), text)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                position = ((1920 - text_width) // 2, (1080 - text_height) // 2)
                draw.text(position, text, fill='black')
            except:
                pass
        
        # Save the image
        if img_ext == 'jpg':
            slide_image = slide_image.convert('RGB')
            slide_image.save(image_path, 'JPEG', quality=95)
        else:
            slide_image.save(image_path, 'PNG')
        
        created_images.append(image_path)
        print(f"Created image: {image_path}")
    
    print(f"\n✓ Successfully created {len(created_images)} image files in '{output_dir}'")
    print("Note: Images show slide content where available. For full rendering, use Windows COM.")
    
    return created_images


def _export_slides_windows_com(input_file, output_dir, image_format='png'):
    """Export slides using Windows PowerPoint COM automation."""
    print("Using Windows PowerPoint COM automation...")
    
    try:
        import win32com.client
    except ImportError:
        raise ImportError(
            "pywin32 not installed. Install it with: pip install pywin32"
        )
    
    # Load presentation to get slide count
    prs = Presentation(input_file)
    total_slides = len(prs.slides)
    input_name = Path(input_file).stem
    
    # Initialize PowerPoint
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    
    try:
        # Open the presentation
        presentation = powerpoint.Presentations.Open(input_file, WithWindow=False)
        
        created_images = []
        img_ext = image_format.lower()
        if img_ext == 'jpeg':
            img_ext = 'jpg'
        
        # Export each slide
        for slide_idx in range(1, total_slides + 1):
            slide = presentation.Slides(slide_idx)
            image_path = os.path.join(output_dir, f"{input_name}_slide_{slide_idx}.{img_ext}")
            
            # Export slide as image
            if img_ext == 'png':
                slide.Export(image_path, "PNG")
            elif img_ext == 'jpg':
                slide.Export(image_path, "JPG")
            
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
    try:
        image_files = export_slides_as_images(input_file, output_images_dir, image_format)
    except Exception as e:
        print(f"Warning: Image generation failed: {e}")
        print("Continuing with PPTX files only...")
        image_files = []
    
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
