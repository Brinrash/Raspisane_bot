"""Bothost-compatible entry point; works locally as well."""
import asyncio
from max_bot import main as run_max_bot

if __name__ == "__main__":
    asyncio.run(run_max_bot())
