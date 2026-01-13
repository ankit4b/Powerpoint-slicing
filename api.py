"""
FastAPI application for PowerPoint Slicer
Provides REST API endpoint to split PowerPoint files into individual slides (PPTX and images)
"""

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Optional
import os
from pptx_slicer import split_pptx, export_slides_as_images, split_pptx_with_images

app = FastAPI(
    title="PowerPoint Slicer API",
    description="API to split PowerPoint presentations into individual slides (PPTX files and images)",
    version="2.0.0"
)


class SliceRequest(BaseModel):
    """Request model for slicing a PowerPoint file"""
    file_path: str
    output_pptx_dir: Optional[str] = None
    output_images_dir: Optional[str] = None
    generate_images: bool = True
    image_format: str = "png"
    
    class Config:
        json_schema_extra = {
            "example": {
                "file_path": "C:/presentations/demo.pptx",
                "output_pptx_dir": "C:/presentations/output/pptx",
                "output_images_dir": "C:/presentations/output/images",
                "generate_images": True,
                "image_format": "png"
            }
        }


class SliceResponse(BaseModel):
    """Response model with details of created files"""
    success: bool
    message: str
    total_slides: int
    pptx_files: List[str]
    image_files: Optional[List[str]] = None


@app.get("/")
def read_root():
    """Root endpoint with API information"""
    return {
        "name": "PowerPoint Slicer API",
        "version": "2.0.0",
        "endpoints": {
            "/slice": "POST - Split a PowerPoint file into individual slides (PPTX + images)",
            "/docs": "GET - Interactive API documentation",
            "/health": "GET - Health check"
        }
    }


@app.get("/health")
def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "PowerPoint Slicer API"}


@app.post("/slice", response_model=SliceResponse)
def slice_powerpoint(request: SliceRequest):
    """
    Split a PowerPoint file into individual slides (PPTX files and images)
    
    - **file_path**: Absolute path to the PowerPoint file (local)
    - **output_pptx_dir**: Optional output directory for PPTX files
    - **output_images_dir**: Optional output directory for image files
    - **generate_images**: Whether to generate images (default: True)
    - **image_format**: Image format - 'png', 'jpg', or 'jpeg' (default: 'png')
    
    Returns list of created PPTX and image file paths
    """
    try:
        # Validate input file exists
        if not os.path.exists(request.file_path):
            raise HTTPException(
                status_code=404,
                detail=f"File not found: {request.file_path}"
            )
        
        # Validate it's a PowerPoint file
        if not request.file_path.lower().endswith(('.pptx', '.ppt')):
            raise HTTPException(
                status_code=400,
                detail="File must be a PowerPoint file (.pptx or .ppt)"
            )
        
        # Validate image format
        if request.image_format.lower() not in ['png', 'jpg', 'jpeg']:
            raise HTTPException(
                status_code=400,
                detail="Image format must be 'png', 'jpg', or 'jpeg'"
            )
        
        # Process the file
        if request.generate_images:
            # Generate both PPTX and images
            result = split_pptx_with_images(
                request.file_path,
                request.output_pptx_dir,
                request.output_images_dir,
                request.image_format
            )
            
            return SliceResponse(
                success=True,
                message=f"Successfully split PowerPoint into {result['total_slides']} slides (PPTX + images)",
                total_slides=result['total_slides'],
                pptx_files=result['pptx_files'],
                image_files=result['image_files']
            )
        else:
            # Generate only PPTX files
            pptx_files = split_pptx(request.file_path, request.output_pptx_dir)
            
            return SliceResponse(
                success=True,
                message=f"Successfully split PowerPoint into {len(pptx_files)} slides (PPTX only)",
                total_slides=len(pptx_files),
                pptx_files=pptx_files,
                image_files=None
            )
        
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing PowerPoint file: {str(e)}"
        )


if __name__ == "__main__":
    import uvicorn
    print("Starting PowerPoint Slicer API...")
    print("API Documentation: http://localhost:8000/docs")
    print("Alternative docs: http://localhost:8000/redoc")
    uvicorn.run(app, host="0.0.0.0", port=8000)
