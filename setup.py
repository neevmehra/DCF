from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = fh.read()

setup(
    name="dcfbot",
    version="1.0.0",
    author="x",
    author_email="x",
    license="GNU GPLv3",
    description="Financial Statement Analyzer CLI",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/",
    py_modules=["dcfbot", "dcfmodel"],
    packages=find_packages(),
    install_requires=requirements,
    python_requires=">=3.7",
    classifiers=[
        "Programming Language :: Python :: 3.10",
        "Operating System :: OS Independent",
    ],
    entry_points={
        "console_scripts": [
            "dcfbot=dcfbot:cli",  # Map "dcfbot" command to the cli function in dcfbot.py
        ],
    }
)