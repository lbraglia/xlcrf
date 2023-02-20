install-from-pipy:
	python3 -m pip install --upgrade xlcrf

install-dev-version:
	pip3 install .

upload:
	rm -rf dist/*
	python3 -m build
	python3 -m twine upload dist/* --verbose

tests: install-dev-version
	cd /tmp && \
	rm -rf *.xlsx  && \
	xlcrf ~/src/pypkg/xlcrf/examples/esempio1.xlsx  && \
	xlcrf ~/src/pypkg/xlcrf/examples/esempio2.xlsx  && \
	libreoffice *.xlsx

wrongexamples: install-dev-version
	cd /tmp && \
	rm -rf *.xlsx  && \
	xlcrf ~/src/pypkg/xlcrf/examples/wrong1.xlsx  && \
	libreoffice *.xlsx
