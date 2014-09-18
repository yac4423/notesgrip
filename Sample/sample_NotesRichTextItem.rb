require 'notesgrip'

ns = Notesgrip::NotesSession.new
db = ns.database("technotes", "call.nsf")

doc = db.CreateDocument("main")

doc['Subject'].text = "RichText TEST"
body = doc.CreateRichTextItem( "body" )
body.text = "Hello, RichText\n"

doc['body'].text += "Add File.\n"
doc['body'].AddFile("./sample_NotesRichTextItem.rb")
doc.save
