from setuptools import setup

from pathlib import Path
this_directory = Path(__file__).parent
long_description = (this_directory / "readme.md").read_text()

setup(name='PyVisualStudioSetupConfiguration',
      version='1.1.0',
      description='Find VisualStudio Setup information',
      long_description = long_description,
      long_description_content_type="text/markdown",
      author='',
      python_requires = ">=3.7",
      py_modules=['PyVisualStudioSetupConfiguration'])
