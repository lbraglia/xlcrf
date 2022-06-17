install:
	pip3 install .

build:
	python3 -m build

upload:
	python3 -m twine upload dist/* --verbose

streamlit:
	streamlit run streamlit.py

# interactive:
# 	inv interactive
