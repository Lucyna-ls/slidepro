# SlidePro

SlidePro is a PowerPoint add-in that helps users create professional presentations by providing intelligent slide recommendations and template suggestions.

## Project Structure

The project consists of two main components:

### SlidePro-Frontend
A PowerPoint VSTO Add-in built with C# that integrates directly into Microsoft PowerPoint, providing:
- Custom ribbon interface
- Task pane for slide recommendations
- Real-time slide analysis and suggestions

### SlidePro-Backend
A FastAPI-based Python backend service that handles:
- Slide analysis and categorization
- Template recommendations using LLM
- JSON processing and metadata extraction

## Requirements

### Frontend
- Visual Studio 2019 or later
- Microsoft Office (PowerPoint)
- .NET Framework 4.7.2

### Backend
- Python 3.8+
- Required packages listed in `requirements.txt`:
  - langchain
  - fastapi
  - uvicorn
  - SQLAlchemy
  - python-dotenv
  - and more...

## Setup

1. Clone both repositories
2. Set up the backend:


```bash
cd SlidePro-Backend
pip install -r requirements.txt
uvicorn src.app.main:app --reload
```

3. Set up the frontend:
   - Open the solution in Visual Studio
   - Restore NuGet packages
   - Build the solution

