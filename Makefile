PYTHON = $(shell test -x bin/python && echo bin/python || echo `which python`)
SETUP  = $(PYTHON) ./setup.py

.PHONY: clean help coverage register sdist upload

help:
	@echo "Please use \`make <target>' where <target> is one or more of"
	@echo "  clean     delete intermediate work product and start fresh"
	@echo "  coverage  run nosetests with coverage"
	@echo "  readme    update README.html from README.rst"
	@echo "  register  update metadata (README.rst) on PyPI"
	@echo "  sdist     generate a source distribution into dist/"
	@echo "  upload    upload distribution tarball to PyPI"

clean:
	find . -type f -name \*.pyc -exec rm {} \;
	rm -rf dist .coverage .DS_Store MANIFEST

coverage:
	nosetests --with-coverage --cover-package=docx --cover-erase

readme:
	rst2html README.rst >README.html
	open README.html

register:
	$(SETUP) register

sdist:
	$(SETUP) sdist

upload:
	$(SETUP) sdist upload
