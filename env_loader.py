from pathlib import Path

from dotenv import load_dotenv


ROOT_DIR = Path(__file__).resolve().parent
ENV_PATH = ROOT_DIR / ".env"


def load_root_env(override: bool = True) -> bool:
    """Load environment variables from the project root .env file."""
    return load_dotenv(dotenv_path=ENV_PATH, override=override)
