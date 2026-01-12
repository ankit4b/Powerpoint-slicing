# PowerPoint Slicer API

FastAPI server has been created! Here's how to use it:

## Starting the Server

```bash
# Using the virtual environment
D:/Python/pptx-slicing/.venv/Scripts/python.exe api.py

# Or if virtual environment is activated
python api.py
```

The server will start at: **http://localhost:8000**

## API Endpoints

### 1. **POST /slice** - Slice a PowerPoint file

**Request Body:**

```json
{
  "file_path": "D:/presentations/demo.pptx",
  "output_dir": "D:/presentations/output" // optional
}
```

**Response:**

```json
{
  "success": true,
  "message": "Successfully split PowerPoint into 10 slides",
  "total_slides": 10,
  "output_files": [
    "D:/presentations/output/demo_slide_1.pptx",
    "D:/presentations/output/demo_slide_2.pptx",
    ...
  ]
}
```

### 2. **GET /health** - Health check

### 3. **GET /** - API information

## Testing the API

### Using curl:

```bash
curl -X POST "http://localhost:8000/slice" \
  -H "Content-Type: application/json" \
  -d "{\"file_path\": \"D:/Python/pptx-slicing/demo.pptx\"}"
```

### Using Python test script:

```bash
python test_api.py
```

### Using Interactive Documentation:

Open in browser: **http://localhost:8000/docs**

## Example with Python requests:

```python
import requests

response = requests.post(
    "http://localhost:8000/slice",
    json={
        "file_path": r"C:\presentations\myfile.pptx",
        "output_dir": r"C:\presentations\sliced"
    }
)

result = response.json()
print(f"Created {result['total_slides']} files")
for file in result['output_files']:
    print(f"  - {file}")
```
