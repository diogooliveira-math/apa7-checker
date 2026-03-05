from setuptools import setup, find_packages

with open("README.md", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="apa7-checker",
    version="1.0.0",
    author="",
    author_email="",
    description=(
        "Automated APA 7 reference checker for Word (.docx) documents. "
        "Produces JSON, HTML, and Word reordering reports."
    ),
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/your-username/apa7-checker",
    packages=find_packages(exclude=["tests*"]),
    install_requires=[
        "python-docx>=0.8.11",
    ],
    extras_require={
        "pdf": ["PyMuPDF"],
        "dev": [
            "pytest>=7.0",
            "pytest-cov",
        ],
    },
    entry_points={
        "console_scripts": [
            "apa7-check=apa7_checker.__main__:main",
        ],
    },
    python_requires=">=3.9",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Science/Research",
        "Intended Audience :: Education",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Text Processing :: Linguistic",
        "Topic :: Utilities",
        "Operating System :: OS Independent",
    ],
    keywords="apa apa7 references citations checker word docx academic",
)
