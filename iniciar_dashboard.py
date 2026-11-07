"""Inicia localmente el dashboard de planificación."""

import os

import uvicorn


if __name__ == "__main__":
    uvicorn.run(
        "dashboard:app",
        host=os.getenv("TELETON_HOST", "127.0.0.1"),
        port=int(os.getenv("TELETON_PORT", "8000")),
        reload=False,
    )
