Q: 
Was hat es mit den VB-Projekten und der Projektgruppe 
in diesem Ordner auf sich?

A:
* In diesem Ordner sind zwei VB-Projekte:
  - MBOFastGeo.vbp,
  - MBOFGbas.vbp
  und eine VB-Projektgruppe:
  - FGGroup.vbg,
  wobei die Projektgruppe die zwei oben genannten Projekte 
  enthält.

* (AXDll) Das VB-Projekt MBOFastGeo.vbp enthält den Sourcecode, 
  also alle Klassen und Module, für die ActiveXdll: 
  - MBOFastGeo.dll,
  die sich nach dem Kompilieren ein Verzeichnis darüber 
  befindet.
  im einzelnen sind das die Dateien:
  Module:
  - ModConsts.bas
  - ModDelphi.bas
  Klassenmodule:
  - FastGeo.cls (der eigentliche Grafikkern als Klasse) 
  - FGTypes.cls
  - TBaseConvexHull.cls
  - TConvexHull2D.cls
  - TOrderedPolygon2D.cls
  - TOrderedPolygon3D.cls
  > wobei die Klasse FastGeo.cls das Attribut hat:
    - Instancing: 6 - GlobalMultiUse  
    d.h. bei Verwendung der dll in einem Projekt muß kein
    Objekt der Klasse FastGeo.cls angelegt werden, alle 
    Funktionen sind vorhanden, wie man es sonst von einem 
    Modul her gewohnt ist. Dies hat den Vorteil, daß die 
    ActiveXdll mit der bas+tlb austasuchbar ist, und 
    umgekehrt.
  > die Klasse FGTypes enthält alle UDTypes als Public Member 
    (dies ist nur in einer Klasse in einer ActiveXdll möglich)
    da die Klasse FGTypes sonst keine eigene Funktionalität 
    besitzt, ist sie mit dem Attribut:
    - Instancing: 2 - PublicNotCreatable
    belegt.
  > die anderen Klassen müssen im Prinzip nicht verwendet 
    werden, da die Funktionalität über Funktionen in der 
    FastGeo.cls bereitgestellt wird. Die Klassen haben demnach
    das Attribut:
    - Instancing: 1 - Private      
  Das Projekt benötigt nur die obengenannten Dateien und 
  ansonsten keine Verweise oder Komponenten.
  Es sollten also keine Probleme beim Kompilieren auftreten. 

* (bas+tlb) Das VB-Projekt MBOFGbas.vbp enthält den Sourcecode, und
  alle Klassen die für die Verwendung der Grafik-Bibliothek ohne 
  die ActiveXdll gedacht ist, also zur Verwendung, als 
  einkompiliertes Modul.
  Im einzelnen sind das die Dateien:
  Module:
  - FastGeo.bas (der eigentliche Grafikkern, als Modul)
  - FGTypes.bas  
  - ModDelphi.bas
  Klassenmodule:
  - TBaseConvexHull.cls
  - TConvexHull2D.cls
  - TOrderedPolygon2D.cls
  - TOrderedPolygon3D.cls
  die letzgenannten vier Klassen und die Datei ModDelphi.bas
  sind die selben Dateien wie im Projekt MBOFastGeo.vbp.
  (deswegen liegen die Projekte im selben Verzeichnis) 
  > Achtung! Die Datei FGTypes.bas enthält nur ein paar 
    Funktionen und keine UD-Types, diese sind auskommentiert.
    Die UDTypes müssen zur Verwendung in einem Standardexe-
    projekt in eine Typelibrary ausgelagert werden.
    Das Projekt hat also einen Verweis auf die Datei 
    - TLBFastGeo.tlb
    Die Datei enthält alle Grafikprimitiven auf die die
    Funktionen abgestimmt sind.
    wo die tlb-Datei liegt bleibt dem Anwender selber überlassen, 
    also in das Windows/System32 Verzeichnis kopieren oder in den 
    Programmordner oder sonstwo. Dort muß die Datei allerdings 
    dann auch registriert werden.
    Sollten beim Start des Projektes komische Phänomene auftauchen,
    ein Absturz von VB gar oder sonstwas, dann liegt das vermutlich 
    daran daß die Datei nicht registriert ist.  
  
  Will man also keine dll verwenden, sondern arbeitet lieber 
  mit der einkompilierten Version, so sind diese Dateien 
  erforderlich um die komplette Funktionalität in einem 
  Standardexe Projekt zur Verfügung zu haben.
  übrigens die ActiveXdll belegt ca. 600kb
  das Modul und tlb nur ein paar kb in der Exe-Datei
  
sollten noch Fragen offen bleiben?

-> Oliver Meyer 
   www.MBO-Ing.com
   oliver.meyer@mbo-ing.com
   olimilo@gmx.net        