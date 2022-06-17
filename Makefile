install-from-pipy:
	python3 -m pip install --upgrade xlcrf

install-dev-version:
	pip3 install .

build:
	rm -rf dist/*
	python3 -m build

upload:
	python3 -m twine upload dist/* --verbose
