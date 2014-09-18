require 'notesgrip'

ns = Notesgrip::NotesSession.new
db = ns.database("technotes", "names.nsf")

# UniversalID
doc = db.CreateDocument("People")
p doc.unid  # UniversalID

# Item Access
item = doc['FullName']  # Create new NotesItem automatically.
richItem = doc.CreateRichTextItem( "Body" )
p doc['Body']  # NotesDocument#[name] return NotesItem or NotesRichTextItem

# GetDocumentByUNID
unid = nil
db.each_document {|doc|
  unid = doc.unid
  break
}
doc = db.GetDocumentByUNID(unid)

# each_item
doc.each_item {|item|
  p item
}

# CopyAllItems
new_doc = db.CreateDocument("People")
new_doc.CopyAllItems(doc)

# Item Copy
new_doc['FullName_Copy'] = doc['FullName']
p new_doc['FullName_Copy']


