#!/usr/bin/env python

"""The setup script."""

from setuptools import setup, find_packages

with open("README.rst") as readme_file:
    readme = readme_file.read()

with open("HISTORY.rst") as history_file:
    history = history_file.read()

requirements = [
    "pywin32",
]

setup_requirements = [
    "pytest-runner",
    "sphinxcontrib-napoleon",
    "sphinx_rtd_theme",
    "sphinx-autoapi",
]

test_requirements = [
    "pytest>=3",
    "psutil>=5",
]

setup(
    author="Josh Coles",
    author_email="josh@colescanada.com",
    python_requires=">=3.5",
    classifiers=[
        "Development Status :: 2 - Pre-Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Natural Language :: English",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
    ],
    description="Solidworks Toolkit for Python",
    install_requires=requirements,
    license="MIT license",
    long_description=readme + "\n\n" + history,
    include_package_data=True,
    keywords="swtoolkit",
    name="swtoolkit",
    packages=find_packages(include=["swtoolkit", "swtoolkit.*"]),
    setup_requires=setup_requirements,
    test_suite="tests",
    tests_require=test_requirements,
    url="https://github.com/Glutenberg/swtoolkit",
    version="0.0.2",
    zip_safe=False,
)
