require 'notesgrip'

ns = Notesgrip::NotesSession.new
p ns.UserName
p ns.AddressBooks # Array of Local names.nsf and Server names.nsf
p ns.CurrentDatabase

# Get database
db = ns.database("myServer", "names.nsf")

# Create NotesName
userName = ns.CreateName("CN=Administrator/O=tech")

# ListUp Databases on server(NotesDbDirectory)

#db_directory = ns.GetDbDirectory( "technotes" )
#db = db_directory.GetFirstDatabase
#while db
#  p db
#  db = db_directory.GetNextDatabase
#end

# ListUp Databases on server
ns.each_database("technotes") {|db|
  p db
}
