.DEFAULT_GOAL := help
.PHONY: help setup dev test lint typecheck format clean

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-15s\033[0m %s\n", $$1, $$2}'

setup: ## Install dependencies in virtual environment
	python -m venv .venv
	.venv/bin/pip install --upgrade pip
	.venv/bin/pip install -e ".[dev]"
	.venv/bin/pip install git+https://github.com/JeffExec/bond-master.git

dev: ## Start bond-master API server for development
	.venv/bin/bondmaster serve

test: ## Run test suite
	.venv/bin/pytest tests/ -v

test-cov: ## Run tests with coverage report
	.venv/bin/pytest tests/ -v --cov=bondmaster_excel --cov-report=term-missing

lint: ## Run linter (ruff)
	.venv/bin/ruff check bondmaster_excel/ tests/

lint-fix: ## Run linter and auto-fix issues
	.venv/bin/ruff check --fix bondmaster_excel/ tests/

typecheck: ## Run type checker (mypy)
	.venv/bin/mypy bondmaster_excel/

format: ## Format code (ruff format)
	.venv/bin/ruff format bondmaster_excel/ tests/

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
setup-win: ## Install dependencies (Windows PowerShell)
	python -m venv .venv
	.venv\Scripts\pip install --upgrade pip
	.venv\Scripts\pip install -e ".[dev]"
	.venv\Scripts\pip install git+https://github.com/JeffExec/bond-master.git

dev-win: ## Start API server (Windows)
	.venv\Scripts\bondmaster serve

test-win: ## Run tests (Windows)
	.venv\Scripts\pytest tests\ -v
