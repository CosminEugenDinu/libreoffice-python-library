import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="libreoffice-py-lib",
    version="0.0.1",
    author="Cosmin Eugen Dinu",
    author_email="tcosdev@gmail.com",
    description="Python-UNO-LibreOffice interface for working with documents.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/CosminEugenDinu/libreoffice-python-library/tree/pypi",
    packages=setuptools.find_packages(),
    install_requires=['toml'],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v2 (GPLv2)",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)