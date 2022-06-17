install-from-pipy:
	python3 -m pip install --upgrade xlcrf

install-dev-version:
	pip3 install .

build:
	rm dist/*
	python3 -m build

upload:
	python3 -m twine upload dist/* --verbose



# python3 -m pip install --index-url https://test.pypi.org/simple/ --no-deps example-package-YOUR-USERNAME-HERE
