import os


def test_readme_exists():
    assert os.path.exists('README.md')


def test_requirements_has_pywin32():
    with open('requirements.txt', 'r', encoding='utf-16') as f:
        data = f.read()
    assert 'pywin32' in data
