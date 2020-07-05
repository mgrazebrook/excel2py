import setuptools

with open("README.md", "r") as f:
    long_description = f.read()

setuptools.setup(
    name='ecel2py',
    version='0.0.1',
    author='Michael Grazebrook',
    author_email='excel2py@grazebrook.com',
    description='Convert an excel file into Python code doing the same calculation',
    long_description=long_description,
    long_description_content_type='text/markdown',
    packages=setuptools.find_packages(
        exclude=('tests',),
        include=('excel2py',),
    ),  # TODO: Should I only install the run-time components?
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Development Status :: 2 - Pre-Alpha",
        "Framework :: Pytest",
        "Intended Audience :: Financial and Insurance Industry",
    ],
    python_requires='>=3.6.3',
)
