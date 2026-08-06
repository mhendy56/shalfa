from app.legacy import app
from app.routers.health import router as health_router
app.include_router(health_router)

def create_app():
    return app
