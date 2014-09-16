# convenience makefile to set up the virtualenv

version = 2.7
venv = virtualenv-$(version)

bin/pserve bin/py.test: bin/python bin/pip setup.py
	bin/python setup.py develop

bin/python bin/pip:
	$(venv) .

clean:
	git clean -fXd

.PHONY: clean
