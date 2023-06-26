from setuptools import setup, find_packages

setup(
    name="image_reader",
    version="1.0.0",
    description="Lector comparador de formularios.",
    author="TechAnt",
    packages=find_packages(),
    install_requires=[
        "Pdf2Png",
        "Imagen2Letra",
        "openpyxl"
    ],
    entry_points={
        "console_scripts": [
            "mi_comando = IMAGE_READER"
        ]
    }
)
