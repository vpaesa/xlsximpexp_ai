# Master Makefile

SUBDIRS = opus gemini copilot opus_libxlsxwriter copilot_libxlsxwriter

all:
	@for dir in $(SUBDIRS); do \
		$(MAKE) -C $$dir; \
	done

win64:
	@for dir in $(SUBDIRS); do \
		$(MAKE) -C $$dir win64; \
	done

clean:
	@for dir in $(SUBDIRS); do \
		$(MAKE) -C $$dir clean; \
	done


.PHONY: all win64 clean $(SUBDIRS)
