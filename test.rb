require 'rubyXL'
require 'mysql'
require 'date'

#This software uses scheduler to daily do an update on cases fetched using mySQL
begin

	#Entramos al archivo de indicadores y definimos la primera hoja.
	#En la primera hoja se encuentra el valor de la ultima fila el cual
	#usaremos para llenar la siguiente fila de datos
	
	indicadores = RubyXL::Parser.parse("indicadores.xlsx")
	worksheet = indicadores[0]
	newLine = worksheet[0][1].value - 1

	con = Mysql.new '192.168.0.208', 'serviciotecnico', 'servicio.2009', 'crm'
	puts con.get_server_info
	rs = con.query( "SELECT COUNT(id) FROM cases WHERE status='Abierto' AND assigned_user_id='764f8d79-c978-a036-bb3c-5627b9915d0a'")
	print "Casos abiertos: " 
	openCases = rs.fetch_row
	puts openCases[0]

	rs = con.query( "SELECT COUNT(id) FROM cases WHERE status='PorRecolectar'")
	print "Equipos en proceso de recoleccion: " 
	recolCases = rs.fetch_row
	puts recolCases[0].to_i

	rs = con.query("SELECT COUNT(id) FROM cases WHERE status='PorDiagnosticar' AND assigned_user_id='764f8d79-c978-a036-bb3c-5627b9915d0a' AND NOT account_id='4195'")
	print "Equipos de clientes por reparar: " 
	repairCases = rs.fetch_row
	puts repairCases

	rs = con.query("SELECT COUNT(id) FROM cases WHERE status='PdteRefacciones' AND type='ServicioTecnico' AND assigned_user_id='764f8d79-c978-a036-bb3c-5627b9915d0a'")
	print "Equipos de clientes pendientes de refacciones: " 
	refacCases = rs.fetch_row
	puts refacCases	

	rs = con.query( "SELECT COUNT(id) FROM cases WHERE status='PdteRefacciones' AND account_id='4195'")
	print "Equipos MAICO pendientes de refacciones: " 
	refacMCases = rs.fetch_row
	puts refacMCases

	rs = con.query( "SELECT COUNT(id) FROM cases WHERE status='PorDiagnosticar' AND account_id='4195'")
	print "Equipos de MAICO por reparar: " 
	repairMCases = rs.fetch_row
	puts repairMCases

	rs = con.query( "SELECT COUNT(id) FROM cases WHERE status='Cotizado'")
	print "Equipos Cotizados: " 
	cotizCases = rs.fetch_row
	puts cotizCases

	rs = con.query( "SELECT COUNT(id) FROM cases WHERE status='PorEnviar'")
	print "Equipos por Enviar: " 
	tosendCases = rs.fetch_row
	puts tosendCases

	rs = con.query( "SELECT COUNT(id) FROM cases WHERE NOT status='Closed' AND type='recoleccion' AND assigned_user_id='764f8d79-c978-a036-bb3c-5627b9915d0a'")
	print "Equipos Prestados: " 
	lendCases = rs.fetch_row
	puts lendCases

	rs = con.query( "SELECT COUNT(id) FROM cases WHERE type='capacitacion' AND NOT status='Closed'")
	print "Instalaciones Pendientes: " 
	instalCases = rs.fetch_row

	rs = con.query(" SELECT * FROM calls WHERE created_by in ('5', '6', '7', '9', '148e8b8c-39b5-0c20-4e41-56588df36fdc', '764f8d79-c978-a036-bb3c-5627b9915d0a', 'a1f68393-c2bf-3359-80d5-58c80c414305') AND direction='Outbound' AND status='Held'")
	print "Llamadas salientes: "
	acc = 0
	rs.each do |row|
		d = DateTime.parse(row[2])
		sum = (Date.today.year * 12 + Date.today.month) - (d.year*12+d.month)
		if sum == 0
			acc = acc + 1
		end
	end
	outCalls = acc
	puts outCalls

	rs = con.query(" SELECT * FROM calls WHERE created_by in ('5', '6', '7', '9', '148e8b8c-39b5-0c20-4e41-56588df36fdc', '764f8d79-c978-a036-bb3c-5627b9915d0a', 'a1f68393-c2bf-3359-80d5-58c80c414305') AND direction='Inbound' AND status='Held'")
	print "Llamadas entrantes: "
	acc = 0
	rs.each do |row|
		d = DateTime.parse(row[2])
		sum = (Date.today.year * 12 + Date.today.month) - (d.year*12+d.month)
		if sum == 0
			acc = acc + 1
		end

	end
	inCalls = acc
	puts inCalls

	rescue Mysql::Error => e
		puts e.errno
		puts e.error

	ensure
		con.close if con

		worksheet.add_cell(newLine, 0, Date.today)
		worksheet.add_cell(newLine,1, openCases[0].to_i)
		worksheet.add_cell(newLine,2, recolCases[0].to_i)
		worksheet.add_cell(newLine,3, repairCases[0].to_i)
		worksheet.add_cell(newLine,4, cotizCases[0].to_i)
		worksheet.add_cell(newLine,5, lendCases[0].to_i)
		worksheet.add_cell(newLine,6, refacCases[0].to_i)
		worksheet.add_cell(newLine,7, outCalls)
		worksheet.add_cell(newLine,8, inCalls)
		worksheet.add_cell(newLine,9, repairMCases[0].to_i)
		worksheet.add_cell(newLine,10, refacMCases[0].to_i)
		newLine = newLine + 2
		worksheet[0][1].change_contents(newLine)
		indicadores.write("indicadores.xlsx")
end