from setuptools import setup

name = 'addresslists'
version = 'v1-dev'

setup(name=name,
    version=version,
    description='...',
    url='https://github.com/witsch/addresslists',
    author='Andreas Zeidler',
    author_email='az at zitc.de',
    classifiers=[
        "Programming Language :: Python",
    ],
    py_modules=['lists'],
    include_package_data=True,
    zip_safe=False,
    install_requires=[
        'setuptools',
        'sqlalchemy == 0.9.9',
        'xlwt',
    ],
    entry_points={
        'console_scripts': [
            'addresslists = lists:main',
        ],
    },
)
