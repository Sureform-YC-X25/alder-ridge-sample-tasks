FROM python:3.12.13-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    ALDER_RIDGE_REQUIRE_SHELL_ISOLATION=1 \
    ALDER_RIDGE_PROJECT_ROOT=/app \
    WORLD_SEED_ROOT=/app/seed \
    WORLD_RUNTIME_ROOT=/workspace \
    WORLD_STATE_ROOT=/state

RUN apt-get update && apt-get install -y --no-install-recommends \
      bubblewrap \
      ca-certificates \
      curl \
      fonts-dejavu-core \
      fonts-liberation \
      libreoffice-calc \
      libreoffice-impress \
      libreoffice-writer \
      nodejs \
      npm \
      poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY pyproject.toml uv.lock /app/
RUN pip install --no-cache-dir uv==0.11.29 \
    && uv lock --check \
    && uv sync --frozen --no-dev --no-editable --no-install-project \
    && rm -rf /root/.cache

COPY runtime /app/runtime
COPY seed /app/seed

RUN uv lock --check \
    && uv sync --frozen --no-dev --no-editable \
    && uv run python runtime/build_agent_runtime.py \
    && rm -rf /root/.cache

ENV PATH="/app/.venv/bin:${PATH}"

RUN mkdir -p /workspace /state && chmod 0777 /workspace /state

EXPOSE 8765

CMD ["hud", "serve", "runtime/env.py", "--host", "0.0.0.0"]
