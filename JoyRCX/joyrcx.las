task	0	; main
               ; setv    30,2,32 ; var30 wordt de constante (2) met waarde 0
                disp 0,0,25                               ;var25 op display
               ; out 2,7 ;motoren aan voor test
               senm 0, 1, 0  ; sensor 0 boolean (switch)???????
               senm 1, 3, 0  ; sensor 1 perc (light)????????????
               sent 0, 1             ;sens 0 switch
               sent 1, 3             ;sens 1 light
Forever:
	setv	31,2,0	; var31 wordt de constante(2) met waarde 0      
                wait	2,40    ;2=const,40*10ms, moet zo lang dat zeker al commando vanuit PC?
                                           ;nb bij timer op 150 ms (beter dan 200 ms)
	chk	0,31,2,2,0,OK	; if 31 (watchdog)=0 dan terug else OK
	plays 5 ;geluid voor test
	dir	1,3     ;omkeren(1) alle motoren(3)  als wel gelijk
                wait          2,60   ;wacht 600 ms
	;dir            1,3	;motoren weer in de goede richting zetten  GAAT FOUT ! ! !
                out 	1,3	;motoren uit
	
OK:
	;plays 3  ;geluid voor test
                ;motor en richting afhankelijk van VBprogramma
                subv          31,2, 32                 ;de waarde van de watchdog constante ervan aftrekken 
                                         ; waarde var31 klopt- 19 mot 3, fwd/4=geen motor/3 mot 3 backw
                chk            0,31,0,2,8,Back	; if 31-watchdog constante >8 dan direction forw else back
                dir              2,7                          ;Forward  ;NB AFHANKELIJK VAN AANSLUITING
                subv          31,2,16                     ;31=31-16 (16=waarde richting)
                jmp Motor  
Back:
                dir              0,7                          ;Backward  
Motor:      chk            0,31,2,2,1,Not1	; if motor <>1 dan not 1
                setv 25,2,1        
                out             1,2                         ;motor 2 uit
                out             2,1                         ;motor 1 aan
                jmp Sens
Not1:       chk            0,31,2,2,2,Not2	; if motor <>2 dan not 2
                setv 25,2,2        
                out             1,1                         ;motor 1 uit             
                out             2,2                         ;motor 2 aan
                jmp Sens
Not2:       chk            0,31,2,2,3,Uit	; if motor <>3 dan Uit (4 = geen motor)
                setv 25,2,3        
                out             2,3                          ;motor 1+2 aan
                jmp Sens
Uit:          out             1,3
Sens:
    	;sensoren lezen en in var21               
                 setv          21,9,1                     ; lichtsensor in v21
                chk           9,0,2,2,1,Suit	; if schak op 1 =1 dan v21=v21+100 anders v21 blijft (Suit)
                sumv        21,2,100                   ;v21=v21+100
Suit:         setv 25,0,21   
                jmp     Forever                         ;OK want watchdog veranderd door VB
	endt









