"""
FastAPI application for PowerPoint Slicer
Provides REST API endpoint to split PowerPoint files into individual slides
"""

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Optional
import os
from pptx_slicer import split_pptx

app = FastAPI(
    title="PowerPoint Slicer API",
    description="API to split PowerPoint presentations into individual slides",
    version="1.0.0"
)


class SliceRequest(BaseModel):
    """Request model for slicing a PowerPoint file"""
    file_path: str
    output_dir: Optional[str] = None
    
    class Config:
        json_schema_extra = {
            "example": {
                "file_path": "C:/presentations/demo.pptx",
                "output_dir": "C:/presentations/output"
            }
        }


class SliceResponse(BaseModel):
    """Response model with details of created files"""
    success: bool
    message: str
    total_slides: int
    output_files: List[str]


@app.get("/")
def read_root():
    """Root endpoint with API information"""
    return {
        "name": "PowerPoint Slicer API",
        "version": "1.0.0",
        "endpoints": {
            "/slice": "POST - Split a PowerPoint file into individual slides",
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
    Split a PowerPoint file into individual slides
    
    - **file_path**: Absolute path to the PowerPoint file (local)
    - **output_dir**: Optional output directory (defaults to 'output' in same directory as input)
    
    Returns list of created file paths
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
        
        # Process the file
        created_files = split_pptx(request.file_path, request.output_dir)
        
        return SliceResponse(
            success=True,
            message=f"Successfully split PowerPoint into {len(created_files)} slides",
            total_slides=len(created_files),
            output_files=created_files
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
