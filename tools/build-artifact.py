import os
os.chdir('/home/user/vba-editor')
s=open('css/style.css').read()
html=open('index.html').read()
body=html[html.index('<header'):html.index('<script src="src/errors.js">')]
js=''
for f in ['src/errors.js','src/tokenizer.js','src/parser.js','src/values.js','src/builtins.js','src/forms.js','src/analyzer.js','src/interpreter.js','src/vba.js','js/help.js','js/designer.js','js/highlight.js','js/intellisense.js','js/app.js']:
    js+='<script>\n'+open(f).read()+'\n</script>\n'
out='<title>VBA-Editor</title>\n<meta name="claude:viewport-fit-sides" content="cover">\n<style>\n'+s+'\nhtml, body { height: 100%; }\n</style>\n'+body+js
sp='/tmp/claude-0/-home-user/e5ceeff3-0ac2-57cb-a46c-e1ba22f0a2b2/scratchpad/'
open(sp+'vba-editor.html','w').write(out)
open(sp+'wrapped.html','w').write('<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"></head><body>'+out+'</body></html>')
