from setuptools import setup, find_packages

setup(
    name="LabfolderPy",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "requests",
        "pandas",
        "numpy",
        "Pillow",
        "matplotlib",
    ],
    author="Daniel Parthier",
    author_email="daniel.parthier@gmail.com",
    description="Python API for Labfolder integration",
)
