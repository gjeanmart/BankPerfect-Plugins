#1.8
import BP

bp_export_folder = BPEval("bp_export_folder")
Ctg = {}
for i, c in enumerate(BP.CategName):
  p = c.find("=")
  Idx = int(c[:p])
  c = c[p+1:].strip()
  Parent = BP.CategName[BP.CategParent[i]]
  p = Parent.find("=")
  PIdx = int(Parent[:p])
  Parent = Parent[p+1:].strip()
  if Idx == PIdx: Ctg[Idx] = [c]
  else: Ctg[Idx] = [c, Parent]

def extract_file_name(path):
    path = path.replace("\\", "/")
    if not "/" in path: return path
    return path.split("/")[-1]

def CtgName(index, SubOnly):
  if not Ctg.has_key(index): return ""
  
  c = Ctg[index]
  IsParent = len(c) == 1
  if SubOnly:
    if IsParent: return ""
    else: return c[0]

  if IsParent: return c[0]
  if CKGroup.Checked: return "%s%s%s" %(c[1], ECtgSep.Text, c[0])
  return c[1]

def Export(Acc):
  dates = [d.replace("-", "/") for d in BP.OperationDate[Acc]]
  modes = []
  for m in BP.OperationMode[Acc]:
    if CKNoNum.Checked and m.find("Chq") == 0: m = "Chèque émis"
    modes.append(m)
  third = [p.replace(Comma, " ") for p in BP.Operationthirdparty[Acc]]
  info = [d.replace(Comma, " ") for d in BP.OperationDetails[Acc]]
  categs = [CtgName(i, 0) for i in BP.OperationCateg[Acc]]
  subs = [CtgName(i, 1) for i in BP.OperationCateg[Acc]]
  vals = [str(v).replace(".", ",") for v in BP.OperationAmount[Acc]]
  mark = [(EMrk1.Text, EMrk2.Text, EMrk3.Text)[m] for m in BP.OperationMark[Acc]]
  cpt = []
  for i in range(len(dates)):
    cpt.append(BP.AccountName[Acc])
  
  fields = []
  if CKDte.Checked: fields.append( (CBDte.ItemIndex, "Date", dates) )
  if CKMod.Checked: fields.append( (CBMod.ItemIndex, "Mode", modes) )
  if CKWho.Checked: fields.append( (CBWho.ItemIndex, "Tiers", third) )
  if CKInf.Checked: fields.append( (CBInf.ItemIndex, "Détails", info) )
  if CKCtg.Checked: fields.append( (CBCtg.ItemIndex, "Catégorie", categs) )
  if CKSub.Checked and not CKGroup.Checked: fields.append( (CBSub.ItemIndex, "Sous-catégorie", subs) )
  if CKVal.Checked: fields.append( (CBVal.ItemIndex, "Montant", vals) )
  if CKMrk.Checked: fields.append( (CBMrk.ItemIndex, "Pointage", mark) )
  if CKCpt.Checked: fields.append( (CBCpt.ItemIndex, "Compte", cpt) )
  fields.sort()

  if len(fields) == 0: return ""
  if CKHead.Checked and Acc == 0: records  = [Comma.join([f[1] for f in fields])]
  else: records = []

  line_count = len(fields[0][2])
  if all: lines_indexes = range(line_count)
  else: lines_indexes = BP.VisibleLines()
  for i in lines_indexes:
    line = []
    for f in fields: line.append(f[2][i])
    records.append(Comma.join(line))

  return records




cs, lb, ck, cb = "csDropDownList", "TLabel", "TCheckBox", "TComboBox"
c = "\n".join([str(i) for i in range(1, 10)])
f = CreateComponent("TForm", None)
f.SetProps(Caption="Paramètres de l'export", Width=520, Height=370, Position="poMainFormCenter", BorderStyle="bsSingle", BorderIcons=["biSystemMenu"])
L0 = CreateComponent(lb, f)
L0.SetProps(Parent=f, Left=20, Top=20, Caption="Exporter le champ")
L0.Font.Style = ["fsBold"]
L1 = CreateComponent(lb, f)
L1.SetProps(Parent=f, Left=130, Top=20, Caption="Ordre")
L1.Font.Style = ["fsBold"]
L2 = CreateComponent(lb, f)
L2.SetProps(Parent=f, Left=220, Top=20, Caption="Options d'export")
L3 = CreateComponent(lb, f)
L3.SetProps(Parent=f, Left=220, Top=52, Caption="Séparer les champs par :")
L4 = CreateComponent(lb, f)
L4.SetProps(Parent=f, Left=220, Top=188, Caption="dans le même champ en les séparant par :")
L5 = CreateComponent(lb, f)
L5.SetProps(Parent=f, Left=220, Top=224, Caption="Symboles spécifiant qu'une opération est :")
L6 = CreateComponent(lb, f)
L6.SetProps(Parent=f, Left=220, Top=244, Caption="non pointée / pointée / rapprochée :")
CBDte = CreateComponent(cb, f)
CBDte.SetProps(Parent=f, Left=132, Top=48, Width=44, Style=cs)
CBDte.Items.Text = c
CBDte.ItemIndex = 0
CKDte = CreateComponent(ck, f)
CKDte.SetProps(Parent=f, Left=20, Top=48, Width=100, Caption="Date", Checked=1)
CKMod = CreateComponent(ck, f)
CKMod.SetProps(Parent=f, Left=20, Top=76, Width=100, Caption="Mode", Checked=1)
CBMod = CreateComponent(cb, f)
CBMod.SetProps(Parent=f, Left=132, Top=76, Width=44, Style=cs)
CBMod.Items.Text = c
CBMod.ItemIndex = 1
CKWho = CreateComponent(ck, f)
CKWho.SetProps(Parent=f, Left=20, Top=104, Width=100, Caption="Tiers", Checked=1)
CBWho = CreateComponent(cb, f)
CBWho.SetProps(Parent=f, Left=132, Top=104, Width=44, ItemIndex=0, Style=cs)
CBWho.Items.Text = c
CBWho.ItemIndex = 2
CKInf = CreateComponent(ck, f)
CKInf.SetProps(Parent=f, Left=20, Top=132, Width=100, Caption="Détails", Checked=1)
CBInf = CreateComponent(cb, f)
CBInf.SetProps(Parent=f, Left=132, Top=132, Width=44, ItemIndex=0, Style=cs)
CBInf.Items.Text = c
CBInf.ItemIndex = 3
CKCtg = CreateComponent(ck, f)
CKCtg.SetProps(Parent=f, Left=20, Top=160, Width=100, Caption="Catégorie", Checked=1)
CBCtg = CreateComponent(cb, f)
CBCtg.SetProps(Parent=f, Left=132, Top=160, Width=44, ItemIndex=0, Style=cs)
CBCtg.Items.Text = c
CBCtg.ItemIndex = 4
CKSub = CreateComponent(ck, f)
CKSub.SetProps(Parent=f, Left=20, Top=188, Width=100, Caption="Sous-catégorie", Checked=1)
CBSub = CreateComponent(cb, f)
CBSub.SetProps(Parent=f, Left=132, Top=188, Width=44, ItemIndex=0, Style=cs)
CBSub.Items.Text = c
CBSub.ItemIndex = 5
CKVal = CreateComponent(ck, f)
CKVal.SetProps(Parent=f, Left=20, Top=216, Width=100, Caption="Montant", Checked=1)
CBVal = CreateComponent(cb, f)
CBVal.SetProps(Parent=f, Left=132, Top=216, Width=44, ItemIndex=0, Style=cs)
CBVal.Items.Text = c
CBVal.ItemIndex = 6
CKMrk = CreateComponent(ck, f)
CKMrk.SetProps(Parent=f, Left=20, Top=244, Width=100, Caption="Pointage", Checked=1)
CBMrk = CreateComponent(cb, f)
CBMrk.SetProps(Parent=f, Left=132, Top=244, Width=44, ItemIndex=0, Style=cs)
CBMrk.Items.Text = c
CBMrk.ItemIndex = 7
CKCpt = CreateComponent(ck, f)
CKCpt.SetProps(Parent=f, Left=20, Top=272, Width=100, Caption="Compte", Checked=1)
CBCpt = CreateComponent(cb, f)
CBCpt.SetProps(Parent=f, Left=132, Top=272, Width=44, ItemIndex=0, Style=cs)
CBCpt.Items.Text = c
CBCpt.ItemIndex = 8
CBSep = CreateComponent(cb, f)
CBSep.SetProps(Parent=f, Left=354, Top=48, Width=56, Style=cs)
CBSep.Items.Text = ";\n,\nTAB\r\n"
CBSep.ItemIndex = 0
CKAll = CreateComponent(ck, f)
CKAll.SetProps(Parent=f, Left=220, Top=90, Width=224, Caption="Exporter tous les comptes", Checked=1)
CKHead = CreateComponent(ck, f)
CKHead.SetProps(Parent=f, Left=220, Top=114, Width=224, Caption="Inclure la ligne d'en-tête", Checked=1)
CKNoNum = CreateComponent(ck, f)
CKNoNum.SetProps(Parent=f, Left=220, Top=138, Width=252, Caption="Remplacer \"Chq XXX\" par \"Chèque émis\"", Checked=1)
CKGroup = CreateComponent(ck, f)
CKGroup.SetProps(Parent=f, Left=220, Top=168, Width=222, Caption="Grouper catégorie et sous-catégorie", Checked=0)
ECtgSep = CreateComponent("TEdit", f)
ECtgSep.SetProps(Parent=f, Left=436, Top=182, Width=54, Text=" > ")
EMrk1 = CreateComponent("TEdit", f)
EMrk1.SetProps(Parent=f, Left=436, Top=238, Width=16, Text="-")
EMrk2 = CreateComponent("TEdit", f)
EMrk2.SetProps(Parent=f, Left=454, Top=238, Width=16, Text="P")
EMrk3 = CreateComponent("TEdit", f)
EMrk3.SetProps(Parent=f, Left=472, Top=238, Width=16, Text="R")
BCnl = CreateComponent("TButton", f)
BCnl.SetProps(Parent=f, Left=171, Top=294, Width=90, Height=25, Caption="Annuler", Cancel=1, ModalResult=2)
BOK = CreateComponent("TButton", f)
BOK.SetProps(Parent=f, Left=271, Top=294, Width=90, Height=25, Caption="Exporter >>", ModalResult=1, Default=1)

if f.ShowModal() == 1:
  Comma  = [";", ",", "\t"][CBSep.ItemIndex]
  all = CKAll.Checked

  if all:
  
    filename = extract_file_name(BP.BankPerfectFileName())[:-3]
    csv_file = "%s%s.csv" %(bp_export_folder, filename)
    file = open(csv_file, "w")
	
    for Acc in range(BP.AccountCount()):
        records = Export(Acc)
        file.write("\n".join(records))
        file.write("\n")
	  
    BP.MsgBox("Les fichiers CSV ont été créés dans le dossier %s" %bp_export_folder, 0)
  else:
    Export(BP.AccountCurrent())