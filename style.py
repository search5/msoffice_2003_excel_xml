__author__ = 'jiho'

class Style:
    def __init__(self, style):
        self.namespaces = {
            'ns': 'urn:schemas-microsoft-com:office:spreadsheet',
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet'
        }

        self.id = style.get("{%s}ID" % self.namespaces['ss'])
        self.underline = ""

        self.style = style

        self.buildFont()

    def getUnderline(self):
        if self.underline != None:
            return self.underline
        else:
            return ""

    def setUnderline(self, underline_val):
        self.underline = underline_val

    @property
    def styleID(self):
        return self.id

    def buildFont(self):
        styleFont = self.style.find("./{%s}Font" % self.namespaces['ns'])

        if styleFont != None:
            styleFontUnderline = styleFont.get("{%s}Underline" % self.namespaces['ss'])
            self.setUnderline(styleFontUnderline)