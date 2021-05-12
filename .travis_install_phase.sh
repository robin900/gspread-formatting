#!/usr/bin/env sh
pip install -U pip
pip uninstall -y six
pip install six>=1.12.0
pip install -r requirements-test.txt
pip install --prefer-binary -e .
if [ $TRAVIS_PYTHON_VERSION =~ '^2.7' ]; then
  pip install rsa<=4.1
fi

