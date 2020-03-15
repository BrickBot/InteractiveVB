task	0	; main
        setv    21,2,0 ; var21 wordt de constante (2) met waarde 0
       ;  out 2,7 ;motoren aan voor test
Forever:
	setv	31,2,0	; var31 wordt de constante(2) met waarde 0      
        wait	2,100    ;2=const,100*10ms, moet zo lang dat zeker al commando vanuit PC 1000?
	chk	0,31,2,0,21,OK	; if 31 (watchdog)=0 dan door else OK
	plays 5 ;geluid voor test
	dir	1,7     ;omkeren(1) alle motoren(7)  als wel gelijk
        wait    2,300   ;wacht 3000 ms
	out 	1,7	;motoren uit
	dir     1,7	;motoren weer in de goede richting zetten
OK:
	; plays 5  ;geluid voor test
        jmp     Forever ;OK want watchdog veranderd door VB
	endt