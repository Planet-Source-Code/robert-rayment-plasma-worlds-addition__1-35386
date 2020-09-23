; Plasma.asm   For NASM  by Robert Rayment 25/5/02

; VB
; ReDim IndexArray(1 To sizex, 1 To sizey)
; ReDim ColArray(1 To sizex, 1 To sizey)
; Redom Colors(0 To 511)
; Randomize Timer
; RandSeed = Rnd * 511
;
; PaletteMaxIndex = 511
; Graininess & scale also depends on palette
; Rapidly changing palettes will be grainier
; StartNoise = 0      (0 - 128) ' graininess
; StartStepsize = 32  (0 - 128) ' scale
;
; Noise = StartNoise
; Stepsize = StartStepsize
;
;'-----------------------------------------------------------------
; Plasma MCode Structure
; Public Type PlasmaStruc
;    sizex As Long           ' 128 - 512
;    sizey As Long           ' 128 - 384
;    PtrIndexArray As Long
;    PtrColArray As Long
;    RandSeed As Long        ' (0-.99999) * 511
;    Noise As Long           ' 0 - 128
;    Stepsize As Long        ' 0 - 128
;    WrapX As Long           ' -1 True, 0 False
;    WrapY As Long           ' -1 True, 0 False
;    PaletteMaxIndex As Long ' 511
;	 PtrColors AS Long
; End Type
; Public PlasmaMCODE As PlasmaStruc
; Public PlasmaArray() As Byte  ' Array to hold machine code
;'-----------------------------------------------------------------
; ptrMC = VarPtr(PlasmaArray(1))
; ptrStruc = VarPtr(PlasmaMCODE.sizex)
; Use:-
; res& = CallWindowProc(ptrMC, ptrStruc, P2, P3, P4)
;                              +8        +12 +16 +20
; where ptrMC is pointer to byte array holding machine code
;'-----------------------------------------------------------------

%macro movab 2      ; name & num of parameters
  push dword %2     ; 2nd param
  pop dword %1      ; 1st param
%endmacro           ; use  movab %1,%2
; Allows eg movab bmW,[ebx+4]

%define sizex           [ebp-4]
%define sizey           [ebp-8]
%define PtrIndexArray   [ebp-12]
%define PtrColArray     [ebp-16]
%define RandSeed        [ebp-20]
%define Noise           [ebp-24]
%define Stepsize        [ebp-28]
%define WrapX           [ebp-32]
%define WrapY           [ebp-36]
%define PaletteMaxIndex [ebp-40]
%define PtrColors       [ebp-44]
;-------------------------------------

%Define StartStepsize   [ebp-52]
%Define RndNoise        [ebp-56]
%Define smax            [ebp-60]
%Define smin            [ebp-64]
%Define Col             [ebp-68]
%Define zmul            [ebp-72]
%Define ix              [ebp-76]
%Define iy              [ebp-80]
%Define kx              [ebp-84]
%Define ky              [ebp-88]
%Define n255            [ebp-92]
%Define index			[ebp-96]
%Define sdiv            [ebp-100]

%Define iyup            [ebp-104]
%Define ixup            [ebp-108]

[BITS 32]

    push ebp
    mov ebp,esp
    sub esp,108
    push edi
    push esi
    push ebx

    ; Copy structure
    mov ebx,[ebp+8]
    
    movab sizex,            [ebx]
    movab sizey,            [ebx+4]
    movab PtrIndexArray,    [ebx+8]
    movab PtrColArray,      [ebx+12]
    movab RandSeed,         [ebx+16]
    movab Noise,            [ebx+20]
    movab Stepsize,         [ebx+24]
    movab WrapX,            [ebx+28]
    movab WrapY,            [ebx+32]
    movab PaletteMaxIndex,  [ebx+36]
	movab PtrColors,		[ebx+40]
;-----------

    mov eax,Stepsize
    mov StartStepsize,eax

	xor eax,eax
	mov RndNoise,eax

    mov eax,255
    mov n255,eax

DO:
    mov ecx,1
FIY:
    mov iy,ecx
    push ecx

    mov ecx,1
FIX:
    mov ix,ecx
    push ecx
	
    ;-----------
    mov eax,Noise
    cmp eax,0
    je NoiseDone    ; RndNoise = 0

    Call Random     ; eax = Rnd * 255 = 0-255

    shl eax,1
    mov ebx,Noise
    mul ebx             ; (Rnd*255)*2*Noise
    mov RndNoise,eax
    
	fild dword RndNoise
    fild dword n255      ; 255,RndNoise
    fdivp st1            ; RndNoise/255
    fistp dword RndNoise
    
	mov eax,RndNoise
	mov ebx,Noise
    sub eax,ebx         ; (Rnd*255)*2*Noise/255 - Noise
    mov ebx,Stepsize
    mul ebx             ; (((Rnd*255)*2*Noise)\255 - Noise)*Stepsize
    mov RndNoise,eax

NoiseDone:
    mov eax,iy
    add eax,Stepsize
    cmp eax,sizey
    jb Tix
        mov eax,WrapY
        cmp eax,0
        je YFalse
        ; Col=IndexArray(ix,1)+RndNoise     
		mov eax,ix
		mov kx,eax
		mov eax,1
		mov ky,eax
		Call near IndexArrayAddr	; edi->addr
		mov bx,[edi]
		movsx eax,bx
		add eax,RndNoise
		mov Col,eax
    
	jmp DoBox
    
	    YFalse:
        ; Col=IndexArray(ix,sizey)+RndNoise
		mov eax,ix
		mov kx,eax
		mov eax,sizey
		mov ky,eax
		Call near IndexArrayAddr	; edi->addr
		mov bx,[edi]
		movsx eax,bx
		add eax,RndNoise
		mov Col,eax

    jmp DoBox

Tix:

    mov eax,ix
    add eax,Stepsize
    cmp eax,sizex
    jb Tstart
        mov eax,WrapX
        cmp eax,0
        je XFalse
        ; Col=IndexArray(1,iy)+RndNoise     
		mov eax,1
		mov kx,eax
		mov eax,iy
		mov ky,eax
		Call near IndexArrayAddr	; edi->addr
		mov bx,[edi]
		movsx eax,bx
		add eax,RndNoise
		mov Col,eax

    jmp DoBox

        XFalse:
        ; Col=IndexArray(sizex,iy)+RndNoise     
		mov eax,sizex
		mov kx,eax
		mov eax,iy
		mov ky,eax
		Call near IndexArrayAddr	; edi->addr
		mov bx,[edi]
		movsx eax,bx
		add eax,RndNoise
		mov Col,eax
    jmp DoBox

Tstart:
    mov eax,Stepsize
    cmp eax,StartStepsize
    jne FourPoint
    ; Col = (Rnd*255)*(2*PaletteMaxIndex)/255-PaletteMaxIndex

    Call Random     	; eax = Rnd * 255 = 0-255

    shl eax,1			; 2*(Rnd*255)
    mov ebx,PaletteMaxIndex
    mul ebx             ; (Rnd*255)*2*PaletteMaxIndex
    mov index,eax
    
	fild dword index
    fild dword n255      ; 255,index
    fdivp st1           ; index/255
    fistp dword index	;((Rnd*255)*2*PaletteMaxIndex)/255
    
	mov eax,index
	mov ebx,PaletteMaxIndex
    sub eax,ebx         ; (Rnd*255)*2*PaletteMaxIndex/255 - PaletteMaxIndex
	mov Col,eax
	jmp DoBox

FourPoint:

    ; Col=IndexArray(ix,iy)
		mov eax,ix
		mov kx,eax
		mov eax,iy
		mov ky,eax
		Call near IndexArrayAddr	; edi->addr (ebx & edx used)
		mov bx,[edi]
		movsx eax,bx
		mov Col,eax
		
		mov edx,Stepsize
		shl edx,1			; edx=2*Stepsize

	; Col = Col + IndexArray(ix+Stepsize,iy)
		add edi,edx 
		mov bx,[edi]
		movsx eax,bx
		add Col,eax

    ; Col = Col + IndexArray(ix+Stepsize,iy+Stepsize)
		push edx
		mov eax,edx
		mov ebx,sizex
		mul ebx			; eax = 2*Stepsize*sizex  (edx changed)
		add edi,eax
		mov bx,[edi]
		movsx eax,bx
		add Col,eax
	
    ; Col = Col + IndexArray(ix,iy+Stepsize)
		pop edx
		sub edi,edx
		mov bx,[edi]
		movsx eax,bx
		add Col,eax

    ; Col = Col\4 + RndNoise
		mov eax,Col
		shr eax,2
		
		add eax,RndNoise
		mov Col,eax

DoBox:
    
	mov eax,Stepsize
    cmp eax,1
    jbe SinglePt
        ; For ky = iy to iy + Stepsize
        ;   If ky <= sizey Then
        ;       For kx = ix to ix + Stepsize
        ;           If kx <= sizex then
        ;               IndexArray(kx,ky) = Col
        ;           End if
        ;       Next kx
        ;   End If
        ; Next ky

    	mov eax,iy
		add eax,Stepsize
		mov iyup,eax

    	mov eax,ix
		add eax,Stepsize
		mov ixup,eax

		mov ecx,iy
	Fy:
		mov ky,ecx
		push ecx

		cmp ecx,sizey
		ja Ny
			mov ecx,ix
		Fx:
			cmp ecx,sizex
			ja Ny

			mov kx,ecx
			Call near IndexArrayAddr	; edi->addr
			mov eax,Col
			mov [edi],ax
			inc ecx
			cmp ecx,ixup		; ecx-ixup
		jbe Fx
	Ny:
		pop ecx
		inc ecx
		cmp ecx,iyup			; ecx-iyup
	jbe Fy

	jmp NIX

SinglePt:
    ; IndexArray(ix,iy) = Col
	mov eax,ix
	mov kx,eax
	mov eax,iy
	mov ky,eax
	Call near IndexArrayAddr	; edi->addr
	mov eax,Col
	mov [edi],ax

    ;-----------
NIX:
    pop ecx
    add ecx,Stepsize
    cmp ecx,sizex
    jbe near FIX
NIY:
    pop ecx
    add ecx,Stepsize
    cmp ecx,sizey
    jbe near FIY

    mov eax,Stepsize
    shr eax,1           ; Stepsize=Stepsize\2
	mov Stepsize,eax
    cmp eax,1
    ja near DO
;-----------------

	; Get Col max/min

	mov eax,10000
	mov smin,eax
	neg eax
	mov smax,eax

	mov ecx,sizey
SAy:
	push ecx
	mov ky,ecx
	mov ecx,sizex
SAx:
	mov kx,ecx
	Call near IndexArrayAddr	; edi->addr
	mov bx,[edi]
	movsx eax,bx
	cmp eax,smin
	jge Tsmax
	mov smin,eax
Tsmax:
	cmp eax,smax
	jle Nax 
	mov smax,eax    
Nax:
	dec ecx
	jnz SAx
	
	pop ecx
	dec ecx
	jnz SAy

	; Calc Stretch color factor zmul

	mov eax,smax
	sub eax,smin
	cmp eax,0
	jg Getzmul
	mov eax,1
Getzmul:
	mov sdiv,eax

	fild dword PaletteMaxIndex	
	fild dword sdiv		; sdiv, PaletteMaxIndex
	fdivp st1			; PaletteMaxIndex/sdiv
	fstp dword zmul		; Real
	;------------

	; Fill ColArray(ix.iy) from Colors(index)
	mov ecx,sizey
CAy:
	push ecx
	mov ky,ecx
	mov ecx,sizex
CAx:
	mov kx,ecx
	Call near IndexArrayAddr	; edi->addr
	mov bx,[edi]
	movsx eax,bx
	mov ebx,smin
	sub eax,ebx
	mov index,eax	; (Col-smin)

	fild dword index
	fld dword zmul
	fmulp st1
	fistp dword index		; Index = (Col-smin)*zmul (0 - 511)
	;............
	mov eax,index

	; Check Index in range
	cmp eax,0
	jge L1
	mov eax,0
L1:
	cmp eax,PaletteMaxIndex
	jle L2
	mov eax,PaletteMaxIndex
L2:
	mov index,eax

	mov [edi],ax			; Put in new Index value
;...................................	

	Call ColorsAddr			; esi->Colors(index)
	Call near ColArrayAddr	; edi->ColArray(kx,ky)
	
	movsd					; [esi] -> [edi]
	
	dec ecx
	jnz CAx
	
	pop ecx
	dec ecx
	jnz CAy

;-----------
GETOUT:
pop ebx
pop esi
pop edi
mov esp, ebp
pop ebp
ret 16
;###########################################
;============================================================
Random:      ; Out: aL & RandSeed = rand(0-255)
    mov eax,011813h     ; 71699 prime 
    imul DWORD RandSeed
    add eax, 0AB209h    ; 700937 prime
    rcr eax,1           ; leaving out gives vertical lines plus
                        ; faint horizontal ones, tartan

    ;----------------------------------------
    jc ok              ; these 2 have little effect
    rol eax,1          ;
ok:                     ;
    ;----------------------------------------
    
    ;----------------------------------------
    ;dec eax            ; these produce vert lines
    ;inc eax            ; & with fsin marble arches
    ;----------------------------------------

    mov RandSeed,eax    ; save seed
    and eax,255

RET
;===========================
IndexArrayAddr:  ; In: NB INTEGER ARRAY, kx,ky,sizex,sizey  Out: edi->addr
	;B = edi + (2* (ky-1) * sizex + 2* (kx-1))
	;B = edi + 2* [(ky-1) * sizex + (kx-1))]
	
	mov edi,PtrIndexArray

	mov eax,ky
	dec eax
	mov ebx,sizex
	mul ebx
	mov ebx,kx
	dec ebx
	add eax,ebx
	shl eax,1	; x2
	add edi,eax
RET
;===========================
ColArrayAddr:    ; In: LONG ARRAY, kx,ky,sizex,sizey  Out: edi->addr
	;B = edi + (4 * (ky-1) * sizex + 4 * (kx-1))
	;B = edi + 4 * [(ky-1) * sizex + (kx-1))]

	mov edi,PtrColArray

	mov eax,ky
	dec eax
	mov ebx,sizex
	mul ebx
	mov ebx,kx
	dec ebx
	add eax,ebx
	shl eax,2		; x4
	add edi,eax
RET
;===========================
ColorsAddr:	     ; In: LONG ARRAY, index (0-511)  Out: esi->addr
	;B = esi + 4 * index 

	mov esi,PtrColors

	mov eax,index
	shl eax,2		; x4
	add esi,eax
RET
;===========================
