from setuptools import setup, find_packages

setup(
    name='SecuRoutingCalcReportApp',
    version='1.0.0',
    description='A CSV to Excel report generator application.',
    author='Your Name',
    author_email='your.email@example.com',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    include_package_data=True,
    install_requires=[
        'pandas',
        'openpyxl',
        'tkinter'
    ],
    entry_points={
        'gui_scripts': [
            'secu_routing_calc_report_app=Secu_Routing_calc_report_app:main'
        ]
    },
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)