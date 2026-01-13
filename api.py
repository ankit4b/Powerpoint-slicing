"""
FastAPI application for PowerPoint Slicer
"""

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, ConfigDict
from typing import List, Optional
import os
from pptx_slicer import split_pptx, export_slides_as_images, split_pptx_with_images

app = FastAPI(
    title="PowerPoint Slicer API",
    description="Split PowerPoint presentations into individual slides",
    version="2.0.0"
)


class SliceRequest(BaseModel):
    """Request model for slicing a PowerPoint file"""
    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "file_path": "D:/presentations/demo.pptx",
                "output_pptx_dir": "D:/presentations/output/pptx",
                "output_images_dir": "D:/presentations/output/images",
                "generate_images": True,
                "image_format": "png"
            }
        }
    )
    
    file_path: str
    output_pptx_dir: Optional[str] = None
    output_images_dir: Optional[str] = None
    generate_images: bool = True
    image_format: str = "png"


class SliceResponse(BaseModel):
    """Response model with details of created files"""
    success: bool
    message: str
    total_slides: int
    pptx_files: List[str]
    image_files: Optional[List[str]] = None


@app.get("/")
def read_root():
    """Root endpoint"""
    return {
        "name": "PowerPoint Slicer API",
        "version": "2.0.0",
        "endpoints": {
            "/slice": "POST - Split PowerPoint into individual slides",
            "/docs": "GET - API documentation",
            "/health": "GET - Health check"
        }
    }


@app.get("/health")
def health_check():
    """Health check"""
    return {"status": "healthy"}


@app.post("/slice", response_model=SliceResponse)
def slice_powerpoint(request: SliceRequest):
    """
    Split PowerPoint into individual slides (PPTX files and/or images)
    """
    try:
        # Validate input file
        if not os.path.exists(request.file_path):
            raise HTTPException(status_code=404, detail=f"File not found: {request.file_path}")
        
        if not request.file_path.lower().endswith(('.pptx', '.ppt')):
            raise HTTPException(status_code=400, detail="File must be a PowerPoint file (.pptx or .ppt)")
        
        if request.image_format.lower() not in ['png', 'jpg', 'jpeg']:
            raise HTTPException(status_code=400, detail="Image format must be 'png', 'jpg', or 'jpeg'")
        
        # Process the file
        if request.generate_images:
            result = split_pptx_with_images(
                request.file_path,
                request.output_pptx_dir,
                request.output_images_dir,
                request.image_format
            )
            
            return SliceResponse(
                success=True,
                message=f"Successfully split PowerPoint into {result['total_slides']} slides",
                total_slides=result['total_slides'],
                pptx_files=result['pptx_files'],
                image_files=result['image_files']
            )
        else:
            pptx_files = split_pptx(request.file_path, request.output_pptx_dir)
            
            return SliceResponse(
                success=True,
                message=f"Successfully split PowerPoint into {len(pptx_files)} slides",
                total_slides=len(pptx_files),
                pptx_files=pptx_files,
                image_files=None
            )
        
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    print("Starting PowerPoint Slicer API...")
    print("API Documentation: http://localhost:8000/docs")
    print("Alternative docs: http://localhost:8000/redoc")
    uvicorn.run(app, host="0.0.0.0", port=8000)
 on http://localhost:8000")
    print("Documentation: http://localhost:8000/docs