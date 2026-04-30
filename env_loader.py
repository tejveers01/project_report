from pathlib import Path
import os

from dotenv import load_dotenv


ROOT_DIR = Path(__file__).resolve().parent
ENV_PATH = ROOT_DIR / ".env"


def _apply_env_aliases() -> None:
    """Backfill legacy/new env variable names so modules can use either form."""
    aliases = {
        "COS_SERVICE_INSTANCE_CRN": "COS_SERVICE_INSTANCE_ID",
        "COS_SERVICE_INSTANCE_ID": "COS_SERVICE_INSTANCE_CRN",
        "COS_BUCKET_NAME": "COS_BUCKET",
        "COS_BUCKET": "COS_BUCKET_NAME",
        "WATSONX_API_KEY": "API_KEY",
        "API_KEY": "WATSONX_API_KEY",
    }

    for target, source in aliases.items():
        if not os.getenv(target) and os.getenv(source):
            os.environ[target] = os.environ[source]

    defaults = {
        "KRA_FOLDER": "Milestone/",
    }

    for key, value in defaults.items():
        if not os.getenv(key):
            os.environ[key] = value


def load_root_env(override: bool = True) -> bool:
    """Load environment variables from the project root .env file."""
    loaded = load_dotenv(dotenv_path=ENV_PATH, override=override)
    _apply_env_aliases()
    return loaded
