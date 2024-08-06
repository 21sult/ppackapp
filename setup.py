from setuptools import setup

setup(
    name='PPack Setup',
    version='1.0',
    packages=find_packages(),
    install_requires=['numpy','pandas','plotly','streamlit','openpyxl','io','streamlit_gsheets']
)