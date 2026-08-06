"""ASGI entry point for Render/Gunicorn/Uvicorn."""
from app.factory import create_app
app=create_app()
