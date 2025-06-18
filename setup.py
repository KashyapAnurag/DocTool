from setuptools import setup, find_packages

setup(
    name='pdf_manhole_extractor', 
    version='0.1.0',           
    packages=find_packages(),   
    
    # Dependencies of required package to run
    install_requires=[
        'pypdf',    # For PDF reading and processing
        'pandas',   # For DataFrame manipulation
        'xlsxwriter', # For writing Excel files
        
    ],
    
    # Metadata 
    author='Anurag Kashyap',
    author_email='anuragkr.kashyap@gmail.com',
    description='A utility to extract structured data from PDF records into Excel.',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/KashyapAnurag/DocTool',
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
        'Topic :: pdf Processing',
        'Topic :: Utilities',
    ],
    python_requires='>=3.7', 
)
