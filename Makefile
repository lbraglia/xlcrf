install:
	pip3 install .

build:
	python3 -m build

test:
	inv

streamlit:
	streamlit run streamlit.py

interactive:
	inv interactive
