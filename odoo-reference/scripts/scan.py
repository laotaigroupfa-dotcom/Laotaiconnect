import ast, os, re, json, sys, glob
root = sys.argv[1]; out = sys.argv[2]
rows = []
for d in sorted(os.listdir(os.path.join(root,'addons'))):
    mf = os.path.join(root,'addons',d,'__manifest__.py')
    if not os.path.exists(mf): continue
    m = ast.literal_eval(open(mf,encoding='utf-8').read())
    new_models=set(); ext_models=set(); methods=0
    for f in glob.glob(os.path.join(root,'addons',d,'models','*.py'))+glob.glob(os.path.join(root,'addons',d,'wizard','*.py')):
        src=open(f,encoding='utf-8').read()
        try: tree=ast.parse(src)
        except Exception: continue
        for node in ast.walk(tree):
            if isinstance(node,ast.ClassDef):
                name=None; inh=[]
                for st in node.body:
                    if isinstance(st,ast.Assign) and len(st.targets)==1 and isinstance(st.targets[0],ast.Name):
                        t=st.targets[0].id
                        try: v=ast.literal_eval(st.value)
                        except Exception: continue
                        if t=='_name' and isinstance(v,str): name=v
                        if t=='_inherit': inh = [v] if isinstance(v,str) else list(v) if isinstance(v,(list,tuple)) else []
                    if isinstance(st,ast.FunctionDef) and not st.name.startswith('_'): methods+=1
                if name and name not in inh: new_models.add(name)
                elif inh: ext_models.update(inh)
    menus=0; actions=0
    for f in glob.glob(os.path.join(root,'addons',d,'views','*.xml')):
        x=open(f,encoding='utf-8').read()
        menus+=len(re.findall(r'<menuitem\b',x))
        actions+=len(re.findall(r'model="ir\.actions\.(?:act_window|server)"',x))
    rows.append(dict(tech=d,name=m.get('name',d),category=m.get('category',''),summary=m.get('summary',''),
        application=bool(m.get('application')),auto_install=bool(m.get('auto_install')),depends=m.get('depends',[]),
        license=m.get('license',''),new_models=sorted(new_models),ext_models=sorted(ext_models-new_models),public_methods=methods,menus=menus,actions=actions,
        description=(m.get('description') or '').strip()[:600]))
json.dump(rows,open(out,'w'),ensure_ascii=False,indent=1)
print(len(rows))
