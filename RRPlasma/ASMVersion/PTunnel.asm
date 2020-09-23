;PTunnel.asm  by Robert Rayment  June 2002

;VB
;' Fill ASM Tunnel Structure
;TunnelStruc.sizex = sizex
;TunnelStruc.sizey = sizey
;TunnelStruc.PtrCoordsX = VarPtr(CoordsX(1, 1)) ' Integer array
;TunnelStruc.PtrCoordsY = VarPtr(CoordsY(1, 1)) ' Integer array
;TunnelStruc.PtrColArray = VarPtr(ColArray(1, 1)) ' Long array
;TunnelStruc.PtrColors = VarPtr(Colors(1))        ' Long array  
;TunnelStruc.PtrIndexArray = VarPtr(IndexArray(1, 1)) ' Integer array
;TunnelStruc.kystart = kystart
;TunnelStruc.kxstart = kxstart
;TunnelStruc.zTrimmer = zTrimmer
;TunnelStruc.NumSmooths = NumSmooths

;
;'ptrMC = VarPtr(TunnelMC(1))
;'ptrStruc = VarPtr(TunnelStruc.SqSize)
;
;' ASM  Cylindrical Projection   CylinProj Opcode = 0
;'res = CallWindowProc(ptrMC, ptrStruc, CylinProj, 0&, 0&)
;   

;' Advance & Spin   ASM  LUT Projection  LUTProj Opcode=1
;'res = CallWindowProc(ptrMC, ptrStruc, LUTProj, 0&, 0&)
;                             8         12          16  20

%macro movab 2      ;name & num of parameters
  push dword %2     ;2nd param
  pop dword %1      ;1st param
%endmacro           ;use  movab %1,%2
;Allows eg  movab bmW,[ebx+4]

%define sizex           [ebp-4]
%define sizey           [ebp-8]
%define PtrCoordsX      [ebp-12]
%define PtrCoordsY      [ebp-16]
%define PtrColArray     [ebp-20] 
%define PtrColors       [ebp-24]
%define PtrIndexArray   [ebp-28]
%define kystart         [ebp-32]
%define kxstart         [ebp-36]
%define zTrimmer		[ebp-40]
%define NumSmooths		[ebp-44]

%define zdx         [ebp-48]
%define zdy         [ebp-52]
%define ixs         [ebp-56]
%define iys         [ebp-60]
%define kx          [ebp-64]
%define ky          [ebp-68]

%Define Col         [ebp-72]
%Define ixdc        [ebp-76]
%Define iydc        [ebp-80]

%define ix          [ebp-84]
%define iy          [ebp-88]

%define zAngle      [ebp-92]   ; Real
%define zRadius     [ebp-96]   ; Real

; For Smoothing
%define im1       [ebp-100]
%define ip1       [ebp-104]
%define savr      [ebp-108]
%define savg      [ebp-112]
%define savb      [ebp-116]
%define n3		  [ebp-120]


[bits 32]

    push ebp
    mov ebp,esp
    sub esp,120
    push edi
    push esi
    push ebx
    

    ; Get structure 
    mov ebx,[ebp+8]

    movab sizex,            [ebx]
    movab sizey,            [ebx+4]
    movab PtrCoordsX,       [ebx+8]
    movab PtrCoordsY,       [ebx+12]
    movab PtrColArray,      [ebx+16]
    movab PtrColors,        [ebx+20]
    movab PtrIndexArray,    [ebx+24] 
    movab kystart,          [ebx+28] 
    movab kxstart,          [ebx+32] 
	movab zTrimmer,			[ebx+36]
	movab NumSmooths,		[ebx+40]

    ; Get opcode
    mov eax,[ebp+12]        ; Opcode&
    
    cmp eax,0
    jne T1
    Call CylinProj
    jmp GETOUT
T1:
    cmp eax,1
    jne T2
    Call LUTProj
T2:
    cmp eax,2
    jne GETOUT
    Call SmoothTunnel

GETOUT:

    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16
;###################################################

CylinProj:        ; Project IndexArray onto a cylinder

	; Set CoordsX&Y to 1
    mov ecx,sizey
COOY1:
    push ecx
    mov ky,ecx
    
    mov ecx,sizex
COOX1:
    mov kx,ecx

    Call CoordsYAddr	; Int
	mov ax,1
	mov [edi],ax		; CoordsY()=1

    Call CoordsXAddr	; Int
	mov ax,1
	mov [edi],ax		; CoordsX()=1

NexCOOX1:
    dec ecx
    jnz near COOX1
NexCOOY1:
    pop ecx
    dec ecx
    jnz near COOY1
	;----------------

	mov eax,sizex
	shr eax,1
	mov ixdc,eax	; sizex\2

	mov eax,sizey
	shr eax,1
	mov iydc,eax	; sizey\2

    mov ecx,sizey
CylY:
    push ecx
    mov ky,ecx
    
	mov eax,ky
	sub eax,iydc
	mov zdy,eax

    mov ecx,sizex
CylX:
    mov kx,ecx
	;----------------
	mov eax,kx
	sub eax,ixdc
	mov zdx,eax

	fild dword zdx
	fild dword zdx
	fmulp st1
	fild dword zdy
	fild dword zdy
	fmulp st1
	faddp st1
	fsqrt
	fstp dword zRadius

	mov eax,sizex
	cmp eax,sizey
	jle XLtY	     ; sizex<=sizey
	;sizex>sizey
	fild dword sizey
	fild dword sizex
	fdivp st1			; st1/st0 sizey/sizex
	fld dword zRadius
	fmulp st1			; (sizey/sizex)*zRadius
	jmp sub2
XLtY:	; sizex<=sizey
	fild dword sizex
	fild dword sizey
	fdivp st1			; st1/st0 sizex/sizey
	fld dword zRadius
	fmulp st1			; (sizex/sizey)*zRadius
sub2:
	fistp dword iys
	mov eax,iys
	sub eax,2
	mov iys,eax

	cmp eax,1
	jl NexCylX
	cmp eax,sizey
	jg NexCylX

	fild dword zdy
	fild dword zdx		; zdx,zdy
	fpatan				; st1/st0  Atan(zdy/zdx)
	fldpi
	faddp st1			; zAngle = zAtan2(zdy,zdx)+pi
	fld dword zTrimmer	; .005
	faddp st1			; zAngle = zAtan2(zdy,zdx)+pi+.005
	
	fild dword sizex
	fldpi
	fldpi
	faddp st1			; 2*pi,sizex
	fdivp st1			; sizex/(2*pi), zAngle
	fmulp st1
	fistp dword ixs

	mov eax,ixs
	cmp eax,1
	jl NexCylX
	cmp eax,sizex
	jg NexCylX

    mov edi,PtrIndexArray
    mov eax,iys
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,ixs
    dec ebx
    add eax,ebx
    shl eax,1       ; *2
    add edi,eax
	xor eax,eax
	mov ax,[edi]	; offset into Colors()
	;mov ixs,eax

;	edi + 4*(ixs-1)    ; Long
    mov edi,PtrColors
    ;mov eax,ixs
    dec eax
    shl eax,2       ; *4
    add edi,eax
    
    mov eax,[edi]
    mov Col,eax
    
    Call ColArrayAddr	; Long
	mov eax,Col
    mov [edi],eax

	; Make LUT
	Call CoordsYAddr
	mov eax,iys
	mov [edi],ax

	Call CoordsXAddr  
	mov eax,ixs
	mov [edi],ax

	;----------------
NexCylX:
    dec ecx
    jnz near CylX
NexCylY:
    pop ecx
    dec ecx
    jnz near CylY

RET
;===================================================

SmoothTunnel:

	mov eax,3
	mov n3,eax

	mov ecx,NumSmooths
NumS:
	push ecx
	;- - - - - - - - - 
	; 2-point Smooth Palette in Colors(0-511) Longs
	mov ecx,511
PalS:
	;------------------

	mov eax,ecx
	dec eax			; im1 = i-1
	cmp eax,0
	jge Cim1
	add eax,511		; 511+im1
Cim1:
	mov im1,eax

	mov eax,ecx
	inc eax			; ip1 = i+1
	cmp eax,511
	jle Cip1
	sub eax,511		; ip1-511
Cip1:
	mov ip1,eax

	mov edi,PtrColors
	mov eax,im1
	shl eax,2		; 4*im1
	add edi,eax
	mov eax,[edi]	; eax =hb- xRGB -lb

	and eax,000FEFEFEh
	shr eax,1
	mov Col,eax

	mov edi,PtrColors
	mov eax,ip1		; ip1
	shl eax,2		; 4*ip1
	add edi,eax
	mov eax,[edi]	; eax =hb- xRGB -lb

	and eax,000FEFEFEh
	shr eax,1
	add Col,eax

	mov edi,PtrColors
	mov eax,ecx		; i
	shl eax,2		; 4*i
	add edi,eax
	mov eax,Col
	mov [edi],eax	; New Colors(i)
	
;------------------
NexPalS:
	dec ecx
	cmp ecx,0
	jge near PalS

	; Smooth IndexArray(sizex,sizey) Integers
	mov ecx,sizex
IndexX:
	push ecx
	mov ix,ecx

	mov ecx,sizey
IndexY:
	mov iy,ecx
	;------------------

	mov eax,iy
	dec eax
	cmp eax,0
	jg StCol
	mov eax,sizey
StCol:
	mov ky,eax
	mov eax,ix
	mov kx,eax
	Call IndexArrayAddr
	xor eax,eax
	mov AL,[edi]
	mov Col,eax

	mov eax,iy
	inc eax
	cmp eax,sizey
	jle Add2Col
	mov eax,sizey
Add2Col:
	mov ky,eax
	Call IndexArrayAddr
	xor eax,eax
	mov AL,[edi]
	add Col,eax

	mov eax,ix
	dec eax
	cmp eax,0
	jg Add3Col
	mov eax,sizex
Add3Col:
	mov kx,eax
	mov eax,iy
	mov ky,eax
	Call IndexArrayAddr
	xor eax,eax
	mov AL,[edi]
	add Col,eax

	mov eax,ix
	inc eax
	cmp eax,sizex
	jle Add4Col
	mov eax,1
Add4Col:
	mov kx,eax
	Call IndexArrayAddr
	xor eax,eax
	mov AL,[edi]
	add Col,eax

	mov eax,ix
	mov kx,eax
	mov eax,iy
	mov ky,eax
	Call IndexArrayAddr
	mov eax,Col
	shr eax,2		; /4
	mov [edi],ax	
		
	;------------------
NexIndY:
	dec ecx
	jnz near IndexY
NexIndX:
	pop ecx
	dec ecx
	jnz near IndexX

	;- - - - - - - - - 
	pop ecx
	dec ecx
	jnz near NumS

RET
;===================================================

LUTProj:
;   '-----------------------------
;   ' VB VB VB LUT Projection
;   For ky = sizey To 1 Step -1
;   For kx = sizex To 1 Step -1
;----------------
;      iy = CoordsY(kx, ky) - kystart
;      If iy < 1 Then iy = sizey + iy
;      ix = CoordsX(kx, ky) - kxstart
;      If ix < 1 Then ix = sizex + ix
;      ColArray(kx, ky) = Colors(IndexArray(ix, iy))
;----------------
;   Next kx
;   Next ky
;   '-----------------------------

;   For ky = sizey To 1 Step -1
;   For kx = sizex To 1 Step -1

    mov ecx,sizey
LUTY:
    push ecx
    mov ky,ecx
    
    mov ecx,sizex
LUTX:
    mov kx,ecx
;----------------

;      iy = CoordsY(kx, ky) - kystart
;      If iy < 1 Then iy = sizey + iy
    
    ; edx=Offset into CoordsY & CoordsX
	mov eax,ky
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,kx
    dec ebx
    add eax,ebx
    shl eax,1       ; *2
	mov edx,eax
    
	mov edi,PtrCoordsY
	add edi,edx

;    Call CoordsYAddr	; Int
    xor eax,eax
	mov ax,[edi]
    sub eax,kystart
    cmp eax,1
    jge iyOK
    add eax,sizey
iyOK:
    mov iy,eax
    
;      ix = CoordsX(kx, ky) - kxstart
;      If ix < 1 Then ix = sizex + ix

    mov edi,PtrCoordsX
	add edi,edx
;    Call CoordsXAddr	; Int
    xor eax,eax
	mov ax,[edi]
    sub eax,kxstart
    cmp eax,1
    jge ixOK
    add eax,sizex
ixOK:
    mov ix,eax
    
;      ColArray(kx, ky) = Colors(IndexArray(ix, iy))

;	edi + 2*[(iy-1) * sizex + (ix-1)]	; Int
    mov edi,PtrIndexArray
    mov eax,iy
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,1       ; *2
    add edi,eax

	xor eax,eax
    mov ax,[edi]
    ;mov ixs,eax

;	edi + 4*(ixs-1)    ; Long
    mov edi,PtrColors
    ;mov eax,ixs
    dec eax
    shl eax,2       ; *4
    add edi,eax
    
    mov eax,[edi]
    mov Col,eax
    
    Call ColArrayAddr	; Long
	mov eax,Col
    mov [edi],eax
    
;----------------
;   Next kx
;   Next ky

NexLUTX:
    dec ecx
    jnz near LUTX
NexLUTY:
    pop ecx
    dec ecx
    jnz near LUTY
RET
;===================================================
ColArrayAddr:  ; In: NB Long ARRAY, kx,ky,256x256  Out: edi->addr
    ;B = edi + 4*[(ky-1) * sizex + (kx-1)]
    
    mov edi,PtrColArray

    mov eax,ky
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,kx
    dec ebx
    add eax,ebx
    shl eax,2           ; *4
    add edi,eax
RET
;===================================================
IndexArrayAddr:   ; Addr = edi + 2*[(ky-1) * sizex + (kx-1)]
    
    mov edi,PtrIndexArray
    jmp CommonIntArray

;===================================================
CoordsXAddr:  ; Addr = edi + 2*[(ky-1) * sizex + (kx-1)]
;PtrCoordsX  Integer Array

    mov edi,PtrCoordsX
    jmp CommonIntArray

;===================================================
CoordsYAddr:  ; Addr = edi + 2*[(ky-1) * sizex + (kx-1)]
;PtrCoordsY  Integer Array

    mov edi,PtrCoordsY

CommonIntArray:

    mov eax,ky
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,kx
    dec ebx
    add eax,ebx
    shl eax,1       ; *2
    add edi,eax
RET
;===========================
