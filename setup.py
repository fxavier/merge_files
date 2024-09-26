from setuptools import setup, find_packages

setup(
    name='merge_files',          
    version='1.0.0',                  
    packages=find_packages(),         
    py_modules=['your_script_name'],  
    install_requires=[                
         'openpyxl',
         'pandas',
        'requests',
        'numpy'
    ],
    entry_points={
        'console_scripts': [          
            'merge_files=merge_files:main', 
        ],
    },
    author='Xavier Nhagumbe',
    author_email='xavier_nhagumbe@echomoz.org',
    description='Merge multiple Excel files into one',
    url='https://github.com/fxavier/merge_files.git',  
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',          
)
