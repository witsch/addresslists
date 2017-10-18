# convenience makefile to set up the virtualenv

version = 2.7

bin/pserve bin/py.test: bin/python bin/pip setup.py
	bin/python setup.py develop

bin/python bin/pip:
	virtualenv -p python$(version) .

clean:
	git clean -fXd

.PHONY: clean
