install:
	pip3 install .

build:
	python3 -m build

test:
	inv


interactive:
	inv interactive
