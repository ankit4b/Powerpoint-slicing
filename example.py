# Example: How to use the PowerPoint Slicer

from pptx_slicer import split_pptx

# Example 1: Basic usage with default output directory
if __name__ == "__main__":
    # Replace 'your_presentation.pptx' with your actual file
    input_file = "demo.pptx"
    
    try:
        # This will create individual slides in the 'output' folder
        created_files = split_pptx(input_file)
        
        print(f"\nSuccessfully created {len(created_files)} files:")
        for file in created_files:
            print(f"  - {file}")
            
    except FileNotFoundError:
        print(f"File not found: {input_file}")
        print("Please place a PowerPoint file in this directory and update the filename.")
    except Exception as e:
        print(f"An error occurred: {e}")


# Example 2: Custom output directory
# Uncomment to use:
# created_files = split_pptx("presentation.pptx", output_dir="my_custom_folder")
