install:
	pip3 install .

build:
	python3 -m build

# test:
# 	xlcrf

streamlit:
	streamlit run streamlit.py

# interactive:
# 	inv interactive
