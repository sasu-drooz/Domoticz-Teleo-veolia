#		   Veolia Teleo Plugin
#
#		   Author:	 zaraki673, 2017
#
"""
<plugin key="Teleo" name="Veolia Teleo" author="zaraki673" version="1.0.0">
	<params>
		<param field="Username" label="Login Veolia" width="150px" required="true"/>
		<param field="Password" label="Password Veolia" width="150px" required="true"/>
		<param field="Mode6" label="Debug" width="75px">
			<options>
				<option label="True" value="Debug"/>
				<option label="False" value="Normal"  default="true" />
			</options>
		</param>
	</params>
</plugin>
"""
import Domoticz
import http.cookiejar, urllib 
import datetime
import xlrd
from xlrd.sheet import ctype_text

class URL:
	
	def __init__(self):
		# On active le support des cookies pour urllib
		cj = http.cookiejar.CookieJar()
		self.urlOpener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
	
	def call(self, url, params = None, referer = None, output = None):
		#Domoticz.Log('Calling url')
		data = None if params == None else urllib.parse.urlencode(params).encode("utf-8")
		request = urllib.request.Request(url, data)
		if referer is not None:
			request.add_header('Referer', referer)
		response = self.urlOpener.open(request)
		#Domoticz.Log(" -> %s" % response.getcode())
		if output is not None:
			file = open(output, 'w')
			file.write(response.read())
			file.close()
		return response

class BasePlugin:
	
	lastHeartbeat = datetime.datetime.now()
	enabled = False
	
	def __init__(self):
		#self.var = 123
		return

	def onStart(self):
		Domoticz.Log("onStart called")
		global conso
		if Parameters["Mode6"] == "Debug":
			Domoticz.Debugging(1)
		if (len(Devices) == 0):
			Domoticz.Device(Name="Status", Unit=1, Type=113, Subtype=0, Switchtype=2).Create()
			Domoticz.Log("Devices created.")
		else:
			if (1 in Devices): conso = Devices[1].nValue
		DumpConfigToLog()
		Domoticz.Heartbeat(60)

	def onStop(self):
		Domoticz.Log("onStop called")

	def onConnect(self, Connection, Status, Description):
		Domoticz.Log("onConnect called")

	def onMessage(self, Connection, Data, Status, Extra):
		Domoticz.Log("onMessage called")

	def onCommand(self, Unit, Command, Level, Hue):
		Domoticz.Log("onCommand called for Unit " + str(Unit) + ": Parameter '" + str(Command) + "', Level: " + str(Level))

	def onNotification(self, Name, Subject, Text, Status, Priority, Sound, ImageFile):
		Domoticz.Log("Notification: " + Name + "," + Subject + "," + Text + "," + Status + "," + str(Priority) + "," + Sound + "," + ImageFile)

	def onDisconnect(self, Connection):
		Domoticz.Log("onDisconnect called")

	def onHeartbeat(self):
		Domoticz.Log("onHeartbeat called")
		################## check if more than 1 day before check & update value
		lastHeartbeatDelta = (datetime.datetime.now()-self.lastHeartbeat).total_seconds()
		if (lastHeartbeatDelta > 86400):
			checkveolia()

global _plugin
_plugin = BasePlugin()

def onStart():
    global _plugin
    _plugin.onStart()

def onStop():
    global _plugin
    _plugin.onStop()

def onConnect(Connection, Status, Description):
    global _plugin
    _plugin.onConnect(Connection, Status, Description)

def onMessage(Connection, Data, Status, Extra):
    global _plugin
    _plugin.onMessage(Connection, Data, Status, Extra)

def onCommand(Unit, Command, Level, Hue):
    global _plugin
    _plugin.onCommand(Unit, Command, Level, Hue)

def onNotification(Data):
    global _plugin
    _plugin.onNotification(Data)

def onDisconnect(Connection):
    global _plugin
    _plugin.onDisconnect(Connection)

def onHeartbeat():
    global _plugin
    _plugin.onHeartbeat()

	# Generic helper functions
def DumpConfigToLog():
	for x in Parameters:
		if Parameters[x] != "":
			Domoticz.Debug( "'" + x + "':'" + str(Parameters[x]) + "'")
	Domoticz.Debug("Device count: " + str(len(Devices)))
	for x in Devices:
		Domoticz.Debug("Device:		   " + str(x) + " - " + str(Devices[x]))
		Domoticz.Debug("Device ID:	   '" + str(Devices[x].ID) + "'")
		Domoticz.Debug("Device Name:	 '" + Devices[x].Name + "'")
		Domoticz.Debug("Device nValue:	" + str(Devices[x].nValue))
		Domoticz.Debug("Device sValue:   '" + Devices[x].sValue + "'")
		Domoticz.Debug("Device LastLevel: " + str(Devices[x].LastLevel))
	return
	
def UpdateDevice(Unit, nValue, sValue):
	# Make sure that the Domoticz device still exists (they can be deleted) before updating it 
	if (Unit in Devices):
		if (Devices[Unit].nValue != nValue) or (Devices[Unit].sValue != sValue):
			Devices[Unit].Update(nValue, sValue)
			Domoticz.Log("Update "+str(nValue)+":'"+str(sValue)+"' ("+Devices[Unit].Name+")")
	return

def checkveolia():
	url = URL()		
	urlConnect = 'https://www.service-client.veoliaeau.fr/home.loginAction.do#inside-space'
	urlConso1 = 'https://www.service-client.veoliaeau.fr/home/espace-client/votre-consommation.html'
	urlConso2 = 'https://www.service-client.veoliaeau.fr/home/espace-client/votre-consommation.html?vueConso=historique'
	urlXls = 'https://www.service-client.veoliaeau.fr/home/espace-client/votre-consommation.exportConsommationData.do?vueConso=historique'
	urlDisconnect = 'https://www.service-client.veoliaeau.fr/logout'
	# Connect to Veolia website
	Domoticz.Log('Connection au site Veolia Eau')
	params = {'veolia_username' : Parameters["Username"] ,
		 'veolia_password' : Parameters["Password"],
		 'login' : 'OK'}
	referer = 'https://www.service-client.veoliaeau.fr/home.html'
	url.call(urlConnect, params, referer)
	# Page 'votre consomation'
	#Domoticz.Log('Page de consommation')
	url.call(urlConso1)
	# Page 'votre consomation : historique'
	#Domoticz.Log('Page de consommation : historique')
	url.call(urlConso2)
	# Download XLS file
	Domoticz.Log('Telechargement du fichier')
	response = url.call(urlXls)
	content = response.read()
	# logout
	Domoticz.Log('Deconnection du site Veolia Eau')
	url.call(urlDisconnect)
	file = open('temp.xls', 'wb')
	file.write(content)
	file.close()
	book = xlrd.open_workbook('temp.xls', encoding_override="cp1252")
	sheet = book.sheet_by_index(0)
	last_rows = sheet.nrows
	row = sheet.row(last_rows-1)
	for idx, cell_obj in enumerate(row):
		cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
		if idx == 1:
			UpdateDevice(1,0,cell_obj.value)
