.PHONY: setup test run docker-build docker-run clean sample

# Install python dependencies
setup:
	pip install -r requirements.txt

# Run pytest unit tests
test:
	pytest

# Run the local python classification pipeline (Regex + LLM)
run:
	python src/run_pipeline.py

# Create a sample mock Excel file for testing
sample:
	python src/make_sample.py

# Build the Docker container for automation
docker-build:
	cd automation && docker compose build

# Run the automation container
docker-run:
	cd automation && docker compose up

# Clean python cache files
clean:
	find . -type d -name "__pycache__" -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
