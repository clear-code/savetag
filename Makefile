PACKAGE_NAME = savetag

all: xpi

xpi:
	./makexpi.sh -n $(PACKAGE_NAME) -o

