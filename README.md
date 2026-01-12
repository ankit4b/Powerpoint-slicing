# PowerPoint Slicer

A Python tool that splits a PowerPoint presentation into individual slides, with each slide saved as a separate PowerPoint file.

## Features

- ✨ Splits any PowerPoint (.pptx) file into individual slides
- 📄 Each slide is saved as a separate .pptx file
- 🎨 Preserves formatting, images, and content
- 📁 Organized output in a dedicated folder
- 🚀 Simple command-line interface

## Installation

1. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Split a PowerPoint file into individual slides:

```bash
python pptx_slicer.py presentation.pptx
```

This will create an `output` folder with individual slide files:

- `presentation_slide_1.pptx`
- `presentation_slide_2.pptx`
- `presentation_slide_3.pptx`
- ... and so on

### Custom Output Directory

Specify a custom output directory:

```bash
python pptx_slicer.py presentation.pptx -o my_slides
```

or

```bash
python pptx_slicer.py presentation.pptx --output my_slides
```

### Using as a Python Module

You can also import and use the function in your own Python scripts:

```python
from pptx_slicer import split_pptx

# Split the presentation
created_files = split_pptx("presentation.pptx", output_dir="slides")

# created_files contains list of all generated file paths
for file in created_files:
    print(f"Created: {file}")
```

## Example

If you have a PowerPoint file with 10 slides:

```bash
python pptx_slicer.py myfile.pptx
```

**Output:**

```
Loading presentation: myfile.pptx
Total slides found: 10
Created: output\myfile_slide_1.pptx
Created: output\myfile_slide_2.pptx
Created: output\myfile_slide_3.pptx
Created: output\myfile_slide_4.pptx
Created: output\myfile_slide_5.pptx
Created: output\myfile_slide_6.pptx
Created: output\myfile_slide_7.pptx
Created: output\myfile_slide_8.pptx
Created: output\myfile_slide_9.pptx
Created: output\myfile_slide_10.pptx

✓ Successfully created 10 PowerPoint files in 'output'
```

## Requirements

- Python 3.6 or higher
- python-pptx library

## How It Works

The tool:

1. Loads the input PowerPoint file
2. Iterates through each slide
3. Creates a new PowerPoint presentation for each slide
4. Copies all shapes, text, images, and formatting
5. Saves each slide as an individual .pptx file

## License

MIT License - Feel free to use and modify as needed!
