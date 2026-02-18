.DEFAULT_GOAL := help
.PHONY: help setup dev test lint typecheck format clean

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-15s\033[0m %s\n", $$1, $$2}'

setup: ## Install dependencies with uv
	uv sync --extra dev
	uv add git+https://github.com/JeffExec/bond-master.git

dev: ## Start bond-master API server for development
	uv run bondmaster serve

test: ## Run test suite
	uv run pytest tests/ -v

test-cov: ## Run tests with coverage report
	uv run pytest tests/ -v --cov=bondmaster_excel --cov-report=term-missing

lint: ## Run linter (ruff)
	uv run ruff check bondmaster_excel/ tests/

lint-fix: ## Run linter and auto-fix issues
	uv run ruff check --fix bondmaster_excel/ tests/

typecheck: ## Run type checker (mypy)
	uv run mypy bondmaster_excel/

format: ## Format code (ruff format)
	uv run ruff format bondmaster_excel/ tests/

check: lint typecheck test ## Run all checks (lint, typecheck, test)

clean: ## Remove build artifacts and cache
	rm -rf .venv/
	rm -rf .pytest_cache/
	rm -rf .mypy_cache/
	rm -rf .ruff_cache/
	rm -rf *.egg-info/
	rm -rf dist/
	rm -rf build/
	rm -rf .coverage
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true

# Windows-specific targets (use with PowerShell)
setup-win: ## Install dependencies with uv (Windows PowerShell)
	uv sync --extra dev
	uv add git+https://github.com/JeffExec/bond-master.git

dev-win: ## Start API server (Windows)
	uv run bondmaster serve

test-win: ## Run tests (Windows)
	uv run pytest tests\ -v
