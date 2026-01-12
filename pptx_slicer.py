"""
PowerPoint Slicer - Split a PowerPoint file into individual slides
Each slide will be saved as a separate PowerPoint file.
"""

import os
import shutil
from pathlib import Path
from pptx import Presentation


def split_pptx(input_file, output_dir=None):
    """
    Split a PowerPoint file into individual slides.
    
    Args:
        input_file (str): Path to the input PowerPoint file
        output_dir (str, optional): Directory to save output files. 
                                   Defaults to 'output' folder in current directory.
    
    Returns:
        list: List of paths to the created PowerPoint files
    """
    # Validate input file
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    # Set up output directory
    if output_dir is None:
        output_dir = "output"
    
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


def main():
    """Main function to run the PowerPoint slicer."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Split a PowerPoint file into individual slides"
    )
    parser.add_argument(
        "input_file",
        help="Path to the input PowerPoint file (.pptx)"
    )
    parser.add_argument(
        "-o", "--output",
        default="output",
        help="Output directory for split slides (default: 'output')"
    )
    
    args = parser.parse_args()
    
    try:
        split_pptx(args.input_file, args.output)
    except Exception as e:
        print(f"Error: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
