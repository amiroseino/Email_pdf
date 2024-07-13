from setuptools import setup, find_packages

setup(
    name='pdf_email_app',
    version='0.1',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'PyPDF2',
        'pandas',
        'Pillow',
        'tqdm',
        'pymupdf',
        'customtkinter',
    ],
    entry_points={
        'console_scripts': [
            'pdf-email-app=pdf_email_app.app:main',
        ],
    },
    author='Your Name',
    author_email='your.email@example.com',
    description='A PDF splitter and email sender application using customtkinter',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://your.project.url',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
