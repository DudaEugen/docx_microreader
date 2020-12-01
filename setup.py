from setuptools import find_packages, setup

setup(
    name='docx_microreader',
    packages=find_packages(include=['docx_microreader', 'docx_microreader.*']),
    version='0.1.0',
    description='Read *.docx files and translate those to another formats',
    author='Duda E.V.',
    install_requires=['pillow'],
)
