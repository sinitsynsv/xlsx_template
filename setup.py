import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="xlsx_template",
    version="0.0.1",
    author="Sergei Sinitsyn",
    author_email="sinitsinsv@gmail.com",
    description="Simple xlsx template engine",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/sinitsynsv/xlsx_template",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
