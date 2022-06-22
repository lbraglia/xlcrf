install-from-pipy:
	python3 -m pip install --upgrade xlcrf

install-dev-version:
	pip3 install .

build: 
	rm -rf dist/*
	python3 -m build

upload: build
	python3 -m twine upload dist/* --verbose

tests: install-dev-version
	cd /tmp && \
	rm -rf *.xlsx  && \
	xlcrf ~/src/pypkg/xlcrf/examples/esempio1.xlsx  && \
	xlcrf ~/src/pypkg/xlcrf/examples/esempio2.xlsx  && \
	libreoffice *.xlsx
