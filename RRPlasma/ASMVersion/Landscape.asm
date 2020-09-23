; Landscape.asm   For NASM by Robert Rayment 27/05/02

; VB

;'Public Type LandscapeStruc
;'   sizex As Long           ' 128 - 512
;'   sizey As Long           ' 128 - 384
;'   PtrColArray As Long
;'   dimx As Long
;'   dimy As Long
;'   PtrSmallArray As Long
;'   PtrIndexArray As Long  ' Integer data
;'   LAH As Long
;'   TANG As Long
;'   xeye As Single
;'   yeye As Single
;'   zvp As Single
;'   zdslope As Single
;'   zslopemult As Single
;'   zhtmult  As Single
;'   DISH As Long
;'   Reflect As Long
;'End Type
;'Public LandscapeMCODE As LandscapeStruc
;'Public LandscapeArray() As Byte  ' Array to hold machine code
;'res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)
;                              +8       +12 +16 +20

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
%define PtrIndexArray   [ebp-28]
%define LAH             [ebp-32]
%define TANG            [ebp-36]
%define xeye            [ebp-40]
%define yeye            [ebp-44]
%define zvp             [ebp-48]
%define zdslope         [ebp-52]
%define zslopemult      [ebp-56]
%define zhtmult         [ebp-60]
%define DISH            [ebp-64]


%define zrayang     [ebp-68]
%define xvp         [ebp-72]
%define yvp         [ebp-76]
%define C           [ebp-80]
%define xray        [ebp-84]
%define yray        [ebp-88]
%define zray        [ebp-92]
%define zdx         [ebp-96]
%define zdy         [ebp-100]
%define n360        [ebp-104]
%define zdz         [ebp-108]
%define zvoxscale   [ebp-112]
%define row         [ebp-116]
%define cstep       [ebp-120]
%define ixr         [ebp-124]
%define iyr         [ebp-128]
%define FloorHt     [ebp-132]
%define zHt         [ebp-136]
%define temp        [ebp-140]
%define n2          [ebp-144]


%define izray       [ebp-148]
%define izHt        [ebp-152]

%define Reflect     [ebp-156]

[BITS 32]
push ebp
mov ebp, esp
sub esp,156   
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
    movab PtrIndexArray,    [ebx+24]
    movab LAH,              [ebx+28]
    movab TANG,             [ebx+32]
    movab xeye,             [ebx+36]
    movab yeye,             [ebx+40]
    movab zvp,              [ebx+44]
    movab zdslope,          [ebx+48]
    movab zslopemult,       [ebx+52]
    movab zhtmult,          [ebx+56]
    movab DISH,             [ebx+60]
    movab Reflect,          [ebx+64]

    mov eax,360
    mov n360,eax

    mov eax,2
    mov n2,eax

    mov eax,TANG
    shl eax,1
    mov temp,eax
    fild dword temp     ; 2*TANG
    fldpi
    fmulp st1           ; 2*TANG*pi
    
    mov eax,dimx
    shr eax,1           ; dimx/2    
    mov temp,eax
    
    fild dword temp
    fsubp st1           ; st0-st1 = 2*TANG*pi - dimx/2
    fstp dword zrayang  ; zrayang = (2*TANG*pi) - dimx/2

    mov eax,xeye
    mov xvp,eax
    mov eax,yeye
    mov yvp,eax

    mov ecx,0
ForC:
    mov C,ecx
    push ecx
;------------------
    mov eax,xvp
    mov xray,eax
    mov eax,yvp
    mov yray,eax
    mov eax,zvp
    mov zray,eax

    
    fld dword zrayang
    fild dword C
    faddp st1           ; zrayang+C
    fild dword n360
    fdivp st1           ; st1/st0 = (zrayang+C)/360
    fst dword temp      ; save ((zrayang+C)/360)
    fcos                ; Cos((zrayang+C)/360)
    fstp dword zdy      ; zdy=Cos((zrayang+C)/360)

    fld dword temp      ; ((zrayang+C)/360)
    fsin                ; Sin((zrayang+C)/360)
    fstp dword zdx      ; zdx=Cos((zrayang+C)/360)

    fild dword n2
    fld dword zdslope
    fmulp st1
    fld dword zslopemult
    fmulp st1           ; zdz = 2*zdslope*zslopemult

    mov eax,DISH        ; -1 True, 0 False
    cmp eax,0
    je zdzDone

    fild dword C
    fild dword dimx
    fild dword n2
    fdivp st1           ; st1/st0 = dimx/2
    fsubp st1           ; st1-st0 = C-dimx/2
    fild dword n360
    fdivp st1           ; st1/st0 = (C-dimx/2)/360
    fcos                ; Cos((C-dimx/2)/360)
    fdivp st1           ; zdz/Cos((C-dimx/2)/360)

zdzDone:
    fstp dword zdz

    mov eax,0
    mov zvoxscale,eax   ; zvoxscale=0

    mov eax,1   
    mov row,eax


    mov ecx,0
Forcstep:
    mov cstep,ecx

    ; Get ixr,iyr & check range
    fld dword xray
    fistp dword ixr     ; ixr=xray

    mov eax,ixr
    cmp eax,1
    jge Tixr

	mov ebx,sizex
    add eax,ebx

	jmp Newixr
Tixr:
    cmp eax,sizex
    jle ixrOK

	mov ebx,sizex
	sub eax,ebx

Newixr:
    mov ixr,eax
ixrOK:

    fld dword yray
    fistp dword iyr     ; iyr=yray

    mov eax,iyr
    cmp eax,1
    jge Tiyr

	mov ebx,sizey
    add eax,ebx
    
	jmp Newiyr
Tiyr:
    cmp eax,sizey
    jle iyrOK

	mov ebx,sizey
	sub eax,ebx

Newiyr:
    mov iyr,eax
iyrOK:
;-----------------

    Call IndexArrayAddr ; In ixr,iyr Out: edi->IndexArray(ixr,iyr)

    mov bx,[edi]
    movzx eax,bx
    mov FloorHt,eax

    fild dword FloorHt
    fld dword zhtmult
    fmulp st1
    fst dword zHt
    fistp dword izHt

    fld dword zray
    fistp dword izray

    mov eax,izHt
    cmp eax,izray       ; izHt-izray
    jle near Bumprays
    
    ; izHt > izray
DO:
    ;Call ColArrayAddr  ; Out esi->ColArray(ixr,iyr)

    mov esi,PtrColArray

    mov eax,iyr
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,ixr
    dec ebx
    add eax,ebx
    shl eax,2       ; x4
    add esi,eax

    ;Call SmallArrayAddr    ; In: C, row  Out: edi->SmallArray(C+1,row)

    mov edi,PtrSmallArray
    
    mov eax,row
    dec eax
    mov ebx,dimx
    mul ebx
    mov ebx,C
    ;inc ebx
    ;dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax

    push esi
    movsd               ; [esi]->[edi]
    pop esi

    mov eax,Reflect
    cmp eax,1
    jne AfterReflect
 
    ; Reflect
    mov ebx,row
    
    mov eax,dimy
    sub eax,ebx     ; dimy-row
    inc eax         ; dimy-row+1  = row'
    
    ;Call SmallArrayAddr    ; In: C, row  Out: edi->SmallArray(C+1,row)
    
    mov edi,PtrSmallArray
    
    dec eax         ; row'-1
    mov ebx,dimx
    mul ebx
    mov ebx,C
    ;inc ebx
    ;dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax

    movsd
    
AfterReflect:

    fld dword zdz
    fld dword zdslope
    faddp st1
    fstp dword zdz

    fld dword zray
    fld dword zvoxscale
    faddp st1
    fst dword zray
    fistp dword izray

    mov eax,row
    inc eax
    mov row,eax

    cmp eax,dimy        ; row-dimy
    jle zraytest
    ; row > dimy
    
    pop ecx
    jmp GETOUT


zraytest:
    mov eax,izray
    cmp eax,izHt        ;izray-izHt
    jle near DO
 

Bumprays:
    
    fld dword xray
    fld dword zdx
    faddp st1
    fstp dword xray

    fld dword yray
    fld dword zdy
    faddp st1
    fstp dword yray

    fld dword zray
    fld dword zdz
    faddp st1
    fstp dword zray

    fld dword zvoxscale
    fld dword zdslope
    faddp st1
    fstp dword zvoxscale

Nexcstep:
    inc ecx
    cmp ecx,LAH         ;ecx-LAH
    jle near Forcstep
;------------------
NexC:
    pop ecx
    inc ecx
    cmp ecx,dimx
    jl near ForC

;-----------
GETOUT:
pop ebx
pop edi
pop esi
mov esp, ebp
pop ebp
ret 16
;#######################################
IndexArrayAddr:  ; In: NB INTEGER ARRAY, kx,ky,sizex,sizey  Out: edi->addr
    ;B = edi + (2* (iyr-1) * sizex + 2* (ixr-1))
    ;B = edi + 2* [(iyr-1) * sizex + (ixr-1))]
    
    mov edi,PtrIndexArray

    mov eax,iyr
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,ixr
    dec ebx
    add eax,ebx
    shl eax,1   ; x2
    add edi,eax
RET
;===========================
ColArrayAddr:    ; In: LONG ARRAY, kx,ky,sizex,sizey  Out: esi->addr
    ;B = edi + (4 * (iyr-1) * sizex + 4 * (ixr-1))
    ;B = edi + 4 * [(iyr-1) * sizex + (ixr-1))]

    mov esi,PtrColArray

    mov eax,iyr
    dec eax
    mov ebx,sizex
    mul ebx
    mov ebx,ixr
    dec ebx
    add eax,ebx
    shl eax,2       ; x4
    add esi,eax
RET
;===========================
SmallArrayAddr:  ; SmallArray(C+1,row), dimy,dimx

    mov edi,PtrSmallArray
    
    mov eax,row
    dec eax
    mov ebx,dimx
    mul ebx
    mov ebx,C
    ;inc ebx
    ;dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax
RET
;===========================

