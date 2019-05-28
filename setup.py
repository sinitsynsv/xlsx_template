import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

with open("requirements.txt") as f:
    requirements = f.read().splitlines()
    requirements = [r for r in requirements if "-i https://pypi.org/simple" not in r]

setuptools.setup(
    name="xlsx_template",
    version="0.0.1",
    author="Sergei Sinitsyn",
    author_email="sinitsinsv@gmail.com",
    description="Simple xlsx template engine",
    include_package_data=True,
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/sinitsynsv/xlsx_template",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    install_requires=requirements,
)
