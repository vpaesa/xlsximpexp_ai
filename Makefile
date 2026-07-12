# Master Makefile

SUBDIRS = chatgpt chatgpt_libxlsxwriter gemini gemini_libxlsxwriter opus opus_libxlsxwriter opus_noexpat copilot copilot_libxlsxwriter

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

