import setuptools

with open("README.md", "r") as f:
    long_description = f.read()

setuptools.setup(
    name="python-spreadsheet",
    version="0.0.1",
    author="Jeffrey Morgan",
    description="A wrapper around the openpyxl Workbook that provides a typewriter style write-and-advance metaphor.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/DrJeffreyMorgan/python-spreadsheet.git",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
)
