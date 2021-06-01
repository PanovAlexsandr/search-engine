import setuptools
with open(r'D:\Семестр 2\find_word\README.md', 'r', encoding='utf-8') as fh:
	long_description = fh.read()

setuptools.setup(
	name='find_word',
	version='1.2',
	author='Alexsus',
	author_email='aleksandr.panov2000@yandex.ru',
	description='search by word',
	long_description=long_description,
	long_description_content_type='text/markdown',
	url='https://github.com/PanovAlexsandr',
	packages=['find_word'],
	classifiers=[
		"Programming Language :: Python :: 3",
		"License :: OSI Approved :: MIT License",
		"Operating System :: OS Independent",
	],
	python_requires='>=3.9',
)