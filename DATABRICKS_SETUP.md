# Databricks Setup Instructions

## Installing Dependencies in Databricks

### 1. Install Python Packages

Add this to a notebook cell:

```python
%pip install python-pptx fastapi uvicorn pydantic pillow pdf2image
```

### 2. Install System Dependencies (LibreOffice & Poppler)

Run this in a notebook cell with `%sh`:

```bash
%sh
# Update package lists
sudo apt-get update

# Install LibreOffice (for PPTX to PDF conversion)
sudo apt-get install -y libreoffice

# Install Poppler utilities (for PDF to image conversion)
sudo apt-get install -y poppler-utils

# Verify installations
libreoffice --version
pdftoppm -v
```

### 3. Alternative: Use Init Script (Recommended for Clusters)

Create an init script file in DBFS:

**Path:** `/dbfs/databricks/scripts/install_libreoffice.sh`

**Content:**

```bash
#!/bin/bash

# Install LibreOffice
apt-get update
apt-get install -y libreoffice poppler-utils

# Verify installations
libreoffice --version
echo "LibreOffice and Poppler installed successfully"
```

**Configure in Cluster Settings:**

1. Go to Cluster Configuration
2. Advanced Options → Init Scripts
3. Add: `dbfs:/databricks/scripts/install_libreoffice.sh`
4. Restart cluster

### 4. Using the PowerPoint Slicer in Databricks

```python
from pptx_slicer import split_pptx_with_images

# Upload your PPTX file to DBFS first
# Example: /dbfs/FileStore/presentations/demo.pptx

input_file = "/dbfs/FileStore/presentations/demo.pptx"
output_pptx_dir = "/dbfs/FileStore/output/pptx"
output_images_dir = "/dbfs/FileStore/output/images"

# Process the file
result = split_pptx_with_images(
    input_file,
    output_pptx_dir,
    output_images_dir,
    image_format='png'
)

print(f"Created {result['total_slides']} slides")
print(f"PPTX files: {len(result['pptx_files'])}")
print(f"Image files: {len(result['image_files'])}")
```

### 5. Running the FastAPI Server in Databricks

```python
import uvicorn
from api import app

# Run the API server
uvicorn.run(app, host="0.0.0.0", port=8000)
```

## Docker Deployment (Alternative)

If deploying as a containerized app:

**Dockerfile:**

```dockerfile
FROM python:3.11-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Copy application files
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Run the API
CMD ["uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000"]
```

## Troubleshooting

### LibreOffice Not Found

```bash
%sh
which libreoffice
# If not found, install:
sudo apt-get install -y libreoffice
```

### Poppler Not Found

```bash
%sh
which pdftoppm
# If not found, install:
sudo apt-get install -y poppler-utils
```

### Permission Issues

```python
import os
os.chmod('/path/to/file', 0o777)
```

## Performance Considerations

- LibreOffice conversion may take 2-5 seconds per slide
- For large presentations (50+ slides), consider parallel processing
- Images at 300 DPI provide good quality but larger file sizes
- Reduce DPI to 150 for smaller files: `convert_from_path(pdf, dpi=150)`
