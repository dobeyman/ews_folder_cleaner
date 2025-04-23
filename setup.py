from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

# Définir directement les dépendances au lieu de lire requirements.txt
requirements = [
    "exchangelib",
    "colorama",
    "rich"
]

setup(
    name="ews-folder-cleaner",
    version="1.0.1",
    author="Jenre",  # Remplacez par votre nom
    author_email="je@n-remy.ovh",  # Remplacez par votre email
    description="A powerful utility for cleaning and managing Exchange mailbox folders via EWS",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/dobeyman/ews_folder_cleaner",  # Remplacez par l'URL de votre dépôt
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "ews-cleaner=ews_folder_cleaner.exchange_cleaner_linux_fixed:main",
        ],
    },
) 