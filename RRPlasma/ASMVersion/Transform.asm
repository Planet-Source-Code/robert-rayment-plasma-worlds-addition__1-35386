; Transform.asm   For NASM  by Robert Rayment  29/5/02

; VB
;' Transform Mcode Structure
;Public Type TransformStruc
;   sizex As Long           ' 128 - 512
;   sizey As Long           ' 128 - 384
;   PtrColArray As Long
;   dimx As Long
;   dimy As Long
;   PtrSmallArray As Long
;   Ptrixdent As Long
;   Ptrzrdsq As Long        ' Single data
;   ixa As Long
;   ixsc As Long
;   zScale As Single
;   zWScale As Single
;   zHscale As Single
;   ShadeNumber As Long
;   PtrzSinTab As Long      ' Single data
;End Type
; Use:-
; res& = CallWindowProc(ptrMC, ptrStruc, P2, P3, P4)
;                              +8        +12 +16 +20

%macro movab 2      ; name & num of parameters
  push dword %2     ; 2nd param
  pop dword %1      ; 1st param
%endmacro           ; use  movab %1,%2
; Allows eg movab bmW,[ebx+4]

%define sizex           [ebp-4]
%define sizey           [ebp-8]
%define PtrColArray     [ebp-12]
%define dimx            [ebp-16]
%define dimy            [ebp-20]
%define PtrSmallArray   [ebp-24]
%define Ptrixdent       [ebp-28]
%define Ptrzrdsq        [ebp-32]
%define ixa              [ebp-36]
%define ixsc            [ebp-40]
%define zScale          [ebp-44]
%define zWScale         [ebp-48]
%define zHscale         [ebp-52]
%define ShadeNumber     [ebp-56]
%define PtrzSinTab      [ebp-60]


%define temp    [ebp-64]

%define iy      [ebp-68]
%define iyyd    [ebp-72]
%define zrd     [ebp-76]

%define ixs     [ebp-80]
%define ixlo    [ebp-84]
%define ixup    [ebp-88]
%define zx      [ebp-92]
%define ixt     [ebp-96]
%define ixxd    [ebp-100]

%define kx      [ebp-104]
%define ky      [ebp-108]

%define Col     [ebp-112]


%define kxx     [ebp-116]
%define kyy     [ebp-120]

%define zpi     [ebp-124]
%define zr2d    [ebp-128]

%define R       [ebp-132]
%define G       [ebp-136]
%define B       [ebp-140]

%define zdivor  [ebp-144]
%define halfpi  [ebp-148]

[BITS 32]
push ebp
mov ebp, esp
sub esp,148
push esi
push edi
push ebx
;-----------
    ; Copy structure
    mov ebx,[ebp+8]
    
    movab sizex,            [ebx]
    movab sizey,            [ebx+4]
    movab PtrColArray,      [ebx+8]
    movab dimx,             [ebx+12]
    movab dimy,             [ebx+16]
    movab PtrSmallArray,    [ebx+20]
    movab Ptrixdent,        [ebx+24]
    movab Ptrzrdsq,         [ebx+28]
    movab ixa,              [ebx+32]
    movab ixsc,             [ebx+36]
    movab zScale,           [ebx+40]
    movab zWScale,          [ebx+44]
    movab zHscale,          [ebx+48]
    movab ShadeNumber,      [ebx+52]
    movab PtrzSinTab,       [ebx+56]
;-----------
    mov eax,180
    mov zr2d,eax
    fild dword zr2d
    fldpi
    fdivp st1           ; rad to deg = zr2d = 180/pi
    fstp dword zr2d

    fldpi
    fild dword dimx
    fdivp st1           ; st1/st0 = zpi = pi/dimx
    fstp dword zpi

    fldpi
	fld1
	fld1
	faddp st1			; 2,pi
	fdivp st1			; st1/st0 = hlafpi = pi/2
	fstp dword halfpi


	mov ecx,sizey
ForIY:
    mov iy,ecx
    push ecx            ;<1<<<<<<<<<<<<<<<<<<

    mov eax,ecx
    dec eax             ; (iy-1)
    mov iyyd,eax
    fild dword iyyd
    fld dword zHscale
    fmulp st1           ; (iy-1)*zHScale
    fistp dword iyyd
    mov eax,iyyd
    inc eax             ; eax = iyyd = 1 + (iy-1)*zHScale

    cmp eax,dimy
    
    jge near NexIY      ; iyyd>dimy

;   jle iyydOK
;    mov eax,dimy
;iyydOK:
    mov iyyd,eax        ; iyyd = 1 + (iy-1)*zHScale OR = dimy

    mov esi,Ptrzrdsq    ; esi->zrdsq(1)
    mov eax,iy
    dec eax
    shl eax,2
    add esi,eax
    mov eax,[esi]
    mov zrd,eax  

    mov esi,Ptrixdent   ; esi-> ixdent(1)
    mov eax,iy
    dec eax
    shl eax,2
    add esi,eax
    mov eax,[esi]
    mov ixlo,eax        ; ixlo = ixdent(iy)
    mov ebx,sizex
    sub ebx,eax             
    mov ixup,ebx        ; ixup = sizex-ixdent(iy)
    
    mov ecx,ixlo

ForIXS:
    mov ixs,ecx
    push ecx            ;<2<<<<<<<<<<<<<<<<<<

    ;--------------  
    fld dword zrd       ; zrdsq(iy)

    fild dword ixa
    fild dword ixs
    fsubp st1           ; st1-st0 = zx = ixa-ixs
    fst dword zx
    fld dword zx
    fmulp st1           ; zx * zx, zrd
    fsubp st1           ; st1-st0 = zrd-zx^2

    fabs                ; get occasional -ve sqr HOW?? 
                        ; Anyway Need fabs

    fsqrt               ; zy = Sqr(zrdsq(iy)-(zx*zx))
    fld dword zx        ; zx,zy  NB zx can be zero!
    fpatan              ; ztheta = ATan(st1/st0) = ATan(zy/zx)              
    fld dword zScale
    fmulp st1           ; zScale*ztheta
    
    fild dword ixsc
    faddp st1           ; ixsc + zScale*ztheta
    fistp dword ixt     ; ixt = ixsc + zScale*ztheta

    mov eax,ixt

    cmp eax,1
    jge Tixt
    add eax,sizex       ; ixt = sizex+ixt
    mov ixt,eax
Tixt:
    cmp eax,sizex
    jbe ixtOK
    sub eax,sizex       ; ixt = ixt - sizex
    mov ixt,eax
    cmp eax,sizex
    jbe ixtOK   
    sub eax,sizex       ; ixt = ixt - sizex
    mov ixt,eax
ixtOK:

    mov eax,ixs
    dec eax
    mov ixxd,eax        ; (ixs-1)
    fild dword ixxd
    fld dword zWScale
    fmulp st1           ; (ixs-1)*zWScale
    fistp dword ixxd
    mov eax,ixxd
    inc eax
    mov ixxd,eax        ; 1+(ixs-1)*zWScale
    cmp eax,dimx
    ja near NexIXS      ; If ixxd>dimx -> next iy

    mov eax,ixt
    mov kx,eax
    mov eax,iy
    mov ky,eax
    Call ColArrayAddr   ; edi -> ColArray(ixt,iy) ERROR
    mov eax,[edi]
    mov Col,eax
    ;==========================
    mov ecx,iyyd

Forkyy: 
    cmp ecx,dimy
    ja near NexIXS
    
    mov kyy,ecx
    push ecx            ;<3<<<<<<<<<<<<<<<<<<
    
    mov ecx,ixxd

Forkxx:
    cmp ecx,dimx
    ja near Nexkxx

    mov kxx,ecx

    cmp ecx,ixxd        ; Need only calc Col once
    jg near ColIn

    ;###############  Start Shade
    
    mov eax,ShadeNumber
    cmp eax,1
    jl near ColIn		; No Shading
						; nb zpi = pi/dimx
						;    halfpi = pi/2
    mov eax,Col         ;  B  G  R  A
    movzx edx,al        ; al ah ax ax
    mov B,edx
    movzx edx,ah
    mov G,edx
    bswap eax           ;  A  R  G  B
    movzx edx,ah        ; al ah ax ax
    mov R,edx

    mov eax,ShadeNumber
    cmp eax,1
	jne TestFor2
                    ; ## Shadenum=1 Left Half shading
    mov eax,kxx
    shr eax,1           ; kxx\2
    mov zdivor,eax
    fild dword zdivor
    fld dword zpi
    fmulp st1
    fsin
    fstp dword zdivor   ; zdivor = Sin(zpi*kxx/2)
    jmp FindCol

TestFor2:
    cmp eax,2
    jne TestFor3
                    ; ## Shadenum=2 Center highlighting
    mov eax,kxx
    mov zdivor,eax
    fild dword zdivor
    fld dword zpi
    fmulp st1
    fsin
    fstp dword zdivor   ; Sin(zpi*kxx)
    jmp FindCol

TestFor3:		; Added  Added Added Added
    cmp eax,3
	jne ColIn
                    ; ## Shadenum=3 Right Half shading
    mov eax,kxx		;		Sin(pi#/2 + zpi*kxx/2)
    shr eax,1           ; kxx\2
    mov zdivor,eax
    fild dword zdivor
    fld dword zpi
    fmulp st1
    fld dword halfpi
	faddp st1
    fsin
	fstp dword zdivor   ; zdivor = Sin(pi/2 + zpi*kxx/2)

FindCol:
    fild dword R
    fld dword zdivor
    fmulp st1
    fistp dword R

    fild dword G
    fld dword zdivor
    fmulp st1
    fistp dword G

    fild dword B
    fld dword zdivor
    fmulp st1
    fistp dword B

    mov eax,R
    shl eax,8
    or eax,G
    shl eax,8
    or eax,B
    mov Col,eax

    ;##############  Shade End

ColIn:
    Call SmallArrayAddr ; edi->SmallArray(kxx,kyy)
    mov eax,Col
    mov [edi],eax

Nexkxx:
    inc ecx             ; kxx+1
    mov eax,ixxd
    inc eax
    cmp eax,ecx         ; (ixxd+1)-kxx
    jge near Forkxx
    
Nexkyy:
    pop ecx
    inc ecx             ; kyy+1
    mov eax,iyyd
    inc eax
    cmp eax,ecx         ; (iyyd+1)-kyy
    jge near Forkyy
    ;==========================

    ;--------------
NexIXS:
    pop ecx
    inc ecx             ; ixs+1
    cmp ecx,ixup        ; (ixs+1)-ixup
    jle near ForIXS
NexIY
    pop ecx
    dec ecx             ; iy-1
    jnz near ForIY


;-----------
GETOUT:
pop ebx
pop edi
pop esi
mov esp, ebp
pop ebp
ret 16
;============================================================
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
    shl eax,2       ; x4
    add edi,eax

RET
;============================================================
SmallArrayAddr:

    mov edi,PtrSmallArray

    mov eax,kyy
    dec eax
    mov ebx,dimx
    mul ebx
    mov ebx,kxx
    dec ebx
    add eax,ebx
    shl eax,2       ; x4
    add edi,eax

RET
;============================================================
