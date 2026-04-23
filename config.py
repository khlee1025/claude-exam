"""
config.py - 환경 변수 및 설정 관리
"""
import os
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# Confluence 설정
CONFLUENCE_URL = os.getenv("CONFLUENCE_URL", "").rstrip("/")
CONFLUENCE_USERNAME = os.getenv("CONFLUENCE_USERNAME", "")
CONFLUENCE_PASSWORD = os.getenv("CONFLUENCE_PASSWORD", "")

# LLM API 설정 (OpenAI 호환)
LLM_API_KEY = os.getenv("LLM_API_KEY", "")
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "")
LLM_MODEL = os.getenv("LLM_MODEL", "gpt-4o")
LLM_MAX_TOKENS = int(os.getenv("LLM_MAX_TOKENS", "4000"))

# DB 설정
DB_PATH = os.getenv("DB_PATH", "confluence_data.db")

# 이미지 저장 폴더
IMAGES_DIR = os.getenv("IMAGES_DIR", "images")


def validate_confluence_config():
    """Confluence 설정 유효성 검사"""
    missing = []
    if not CONFLUENCE_URL:
        missing.append("CONFLUENCE_URL")
    if not CONFLUENCE_USERNAME:
        missing.append("CONFLUENCE_USERNAME")
    if not CONFLUENCE_PASSWORD:
        missing.append("CONFLUENCE_PASSWORD")
    if missing:
        raise ValueError(f"❌ 누락된 Confluence 설정: {', '.join(missing)}\n.env 파일을 확인하세요.")


def validate_llm_config():
    """LLM API 설정 유효성 검사"""
    missing = []
    if not LLM_API_KEY:
        missing.append("LLM_API_KEY")
    if not LLM_BASE_URL:
        missing.append("LLM_BASE_URL")
    if missing:
        raise ValueError(f"❌ 누락된 LLM 설정: {', '.join(missing)}\n.env 파일을 확인하세요.")


# 하위 호환성 alias
validate_anthropic_config = validate_llm_config
