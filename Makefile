install:
	poetry install --no-root

notebook: install
	poetry run jupyter notebook --no-browser
