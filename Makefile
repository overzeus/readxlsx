.PHONY : clean
objs = readxlsx.o miniunz.o ioapi.o unzip.o
Target = libreadxlsx.so.1.0

HOMEDIR = $(WS_HOME)

HOME3RD = $(HOMEDIR)/3rdLib/

INCDIR = $(HOMEDIR)/src/include/

LIBDIR = $(HOMEDIR)/lib/

LDFLAGS = -lz -lxml2 -lz

$(Target): $(objs)
	gcc -g -shared $^ -o $@ -I$(INCDIR)
	ar ru libreadxlsx.a $^
	rm -f $(LIBDIR)/libreadxlsx.a
	cp -f libreadxlsx.so.1.0 $(LIBDIR)/
	cp -f libreadxlsx.a $(LIBDIR)/
	ln -sf  libreadxlsx.so.1.0 $(LIBDIR)/libreadxlsx.so

%.o: %.c
	gcc -I$(INCDIR) -c  -g -fPIC -Wall  $< -o $@ -L$(LIBDIR) $(LDFLAGS)

sample: 
	gcc  -I$(INCDIR) sample.c -L$(LIBDIR) -lreadxlsx $(LDFLAGS) -osample -g

clean:
	rm -rf *.o *.so.* core sample a.out
