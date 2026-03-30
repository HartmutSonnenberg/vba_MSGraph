# vba_MSGraph
enthält alle Module und Klassen um aus VBA heraus die Funktionen userread und sendmail in MSGraph zu verwenden. Voraussetzung ist ein Tenant in Microsoft 365, Microsoft Azure und dort eine registrierte App!

clsEncodeImage.cls
  Base64 encodiertes Umwandeln binärer Daten
clsMSGraph.cls
  alle Properties und Methoden, um für userread und sendmail den Authorisierungsflow zu bedienen, also Autorisierungscode anfordern (Voraussetzung in Avure angemeldeter Benutzer), 
  ein Token und dann den eigentlichen Request

  
