; Effects.asm   For NASM by Robert Rayment 26/05/02

; VB

;'Effects MCode Structure
;Public Type EffectsStruc
;   sizex As Long           ' 128 - 512
;   sizey As Long           ' 128 - 384
;   PtrColArray As Long
;End Type
;   Opcode As Long          ' 0 x, 1 y, 2 x & y - smoothing
;                           ' 3 brighten, 4 darken (Increment=4)
;Public EffectsMCODE As EffectsStruc
;Public EffectsArray() As Byte  ' Array to hold machine code
;'res = CallWindowProc(ptrMC, ptrStruc, 0pcode, 0&, 0&)
;                             +8        +12
; where ptrMC is pointer to byte array holding machine code

%macro movab 2      ; name & num of parameters
  push dword %2     ; 2nd param
  pop dword %1      ; 1st param
%endmacro           ; use  movab %1,%2
; Allows eg movab bmW,[ebx+4]

%Define Opcode     [ebp+12]

%Define sizex           [ebp-4]
%Define sizey           [ebp-8]
%Define PtrColArray     [ebp-12]

%Define ix              [ebp-16]
%Define iy              [ebp-20]
%Define kx              [ebp-24]
%Define ky              [ebp-28]
%Define ixL             [ebp-32]
%Define ixR             [ebp-36]
%Define iyA             [ebp-40]
%Define iyB             [ebp-44]

%Define Increment       [ebp-48]
%Define Mask2           [ebp-52]

%Define ColArraySize    [ebp-56]

; For MMX
%Define lo32 [ebp-60]
%Define hi32 [ebp-64]

%Define Mask4      [ebp-68]
%Define SUMMER     [ebp-72]

[BITS 32]
push ebp
mov ebp, esp
sub esp,72
push esi
push edi
push ebx
;-----------
    ; Copy structure
    mov ebx,[ebp+8]
    
    movab sizex,            [ebx]
    movab sizey,            [ebx+4]
    movab PtrColArray,      [ebx+8]

    mov eax,4
    mov Increment,eax

    mov eax,0FEFEFEFEh
    mov Mask2,eax           ; for 2-point smoothing

    mov eax,0F8F8F8F8h
    mov Mask4,eax           ; for 8-point smoothing

    mov eax,Opcode
    cmp eax,0
    jne T1

    Call XSmooth
    jmp GETOUT
T1:
    cmp eax,1
    jne T2

    Call YSmooth
    jmp GETOUT
T2:
    cmp eax,2
    jne T3

    ;Call XSmooth
    ;Call YSmooth

    Call XYSmoothing

    jmp GETOUT
T3:
    cmp eax,3
    jne T4

    Call Brighten
    jmp GETOUT
T4:
    cmp eax,4
    jne GETOUT

    Call Darken

;-----------
GETOUT:
pop ebx
pop edi
pop esi
mov esp, ebp
pop ebp
ret 16
;################################################

XSmooth:
    mov ecx,sizey

ForY0:
    mov ky,ecx
    push ecx

    mov ecx,sizex
ForX0:

    mov eax,ecx     ; ix
    cmp eax,1
    jne xLOK
    mov eax,sizex
    inc eax
xLOK:
    dec eax         ; ix-1 or sizex
    mov ixL,eax

    mov eax,ecx     ; ix
    cmp eax,sizex
    jne xROK
    mov eax,0
xROK:
    inc eax         ; ix+1 or 1
    mov ixR,eax

    mov eax,ixL
    mov kx,eax
    Call near ColArrayAddr2  ; edi->xL
    mov eax,[edi]   ; FAILS HERE  ky=iy=256  ecx,ix=320 - ixL,kx=319
    and eax,Mask2
    shr eax,1
    push eax

    mov eax,ixR
    mov kx,eax
    Call ColArrayAddr2   ; edi->xR
    mov eax,[edi]
    and eax,Mask2
    shr eax,1
    
    pop ebx
    add eax,ebx
    push eax

    mov kx,ecx          ; ix
    Call ColArrayAddr2  ; edi->ix
    pop eax
    mov [edi],eax
;-----

NexX0;
    dec ecx
    jnz ForX0
NexY0:
    pop ecx
    dec ecx
    jnz ForY0

RET
;================================================

YSmooth:

    mov ecx,sizex

ForX1:
    mov kx,ecx
    push ecx

    mov ecx,sizey
ForY1:

    mov eax,ecx     ; iy
    cmp eax,1
    jne yAOK
    mov eax,sizey
    inc eax
yAOK:
    dec eax         ; iy-1 or sizey
    mov iyA,eax


    mov eax,ecx     ; iy
    cmp eax,sizey
    jne yBOK
    mov eax,0
yBOK:
    inc eax         ; iy+1 or 1
    mov iyB,eax

    mov eax,iyA
    mov ky,eax
    Call near ColArrayAddr2  ; edi->yA
    mov eax,[edi]
    and eax,Mask2
    shr eax,1
    push eax

    mov eax,iyB
    mov ky,eax
    Call ColArrayAddr2   ; edi->yB
    mov eax,[edi]
    and eax,Mask2
    shr eax,1
    
    pop ebx
    add eax,ebx
    push eax

    mov ky,ecx          ; iy
    Call ColArrayAddr2  ; edi->iy
    pop eax
    mov [edi],eax


;-----

NexY1;
    dec ecx
    jnz ForY1
NexX1:
    pop ecx
    dec ecx
    jnz ForX1

RET
;================================================
XYSmoothing:    ; 8-point smoothing

    mov ecx,sizey

ForY8:
    mov iy,ecx
    mov ky,ecx
    push ecx

    mov ecx,sizex
ForX8:
    mov ix,ecx
    mov kx,ecx
    
    Call ColArrayAddr2  ; edi->center x,y
    push edi
    pop esi             ; esi->center x,y
    
    mov eax,ix
    dec eax             ; ix-1
    cmp eax,0
    jg Left_ixm1
    mov eax,sizex
Left_ixm1:
    mov kx,eax          ; kx set at ix-1
    
    mov eax,iy
    mov ky,eax          ; ix-1,iy
    Call ColArrayAddr2  ; edi->ix-1,iy
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    mov SUMMER,eax
    
    mov eax,iy
    dec eax             ; iy-1
    cmp eax,0
    jg Left_iym1
    mov eax,sizey
Left_iym1:
    mov ky,eax
    Call ColArrayAddr2  ; edi->ix-1,iy-1
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    add SUMMER,eax
    
    mov eax,iy
    inc eax             ; iy+1
    cmp eax,sizey
    jle Left_iyp1
    mov eax,1
Left_iyp1:
    mov ky,eax
    Call ColArrayAddr2  ; edi->ix-1,iy+1
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    add SUMMER,eax
    ;.. left col done ..............

    mov eax,ix
    mov kx,eax          ; kx reset

    mov eax,iy
    dec eax             ; iy-1
    cmp eax,0
    jg Cen_iym1
    mov eax,sizey
Cen_iym1:
    mov ky,eax
    Call ColArrayAddr2  ; edi->ix-1,iy-1
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    add SUMMER,eax
    
    mov eax,iy
    inc eax             ; iy+1
    cmp eax,sizey
    jle Cen_iyp1
    mov eax,1
Cen_iyp1:
    mov ky,eax
    Call ColArrayAddr2  ; edi->ix-1,iy+1
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    add SUMMER,eax
    ;.. center col done

    mov eax,ix
    inc eax             ; ix+1
    cmp eax,sizex
    jle Right_ixm1
    mov eax,1
Right_ixm1:
    mov kx,eax          ; kx set at ix+1
    
    mov eax,iy
    mov ky,eax          ; ix+1,iy
    Call ColArrayAddr2  ; edi->ix-1,iy
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    add SUMMER,eax
    
    mov eax,iy
    dec eax             ; iy-1
    cmp eax,0
    jg Right_iym1
    mov eax,sizey
Right_iym1:
    mov ky,eax
    Call ColArrayAddr2  ; edi->ix+1,iy-1
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    add SUMMER,eax
    
    mov eax,iy
    inc eax             ; iy+1
    cmp eax,sizey
    jle Right_iyp1
    mov eax,1
Right_iyp1:
    mov ky,eax
    Call ColArrayAddr2  ; edi->ix+1,iy+1
    mov eax,[edi]
    and eax,Mask4
    shr eax,3
    add SUMMER,eax
    ;.. Right col done ..............

    ; Set color at ColArray(ix,iy)

    mov eax,SUMMER
    mov [esi],eax
    
;-----
NexX8;
    dec ecx
    jnz near ForX8
NexY8:
    pop ecx
    dec ecx
    jnz near ForY8

RET
;================================================
Brighten:

; Using MMX

    mov edi,PtrColArray ; edi-> ColArray(1,1)
    mov eax,sizey
    mov ebx,sizex
    mul ebx
    mov ColArraySize,eax    ; Num 4-byte chunks

    xor eax,eax
    mov eax,Increment
    shl eax,8           ;00 00 In 00
                        ;A  R  G  B
    mov aL,Increment
    shl eax,8           ;00 In In 00
                        ;A  R  G  B
    mov aL,Increment    ;00 In In In
                        ;A  R  G  B
    
    mov lo32,eax
    mov hi32,eax
    movq mm0,hi32       ; 8 increments to add to bytes
    
    mov ecx,ColArraySize
    shr ecx,1           ; Num 8-byte chunks
    mov ebx,8           ; For shifting edi 8 bytes at a time
LB:
    movq mm1,[edi]
    paddusb mm1,mm0     ; add with saturation at 255 for each byte
    movq [edi],mm1
    add edi,ebx
    dec ecx
    jnz LB
    emms

RET
;================================================

Darken:

; Using MMX

    mov edi,PtrColArray ; edi-> ColArray(1,1)
    mov eax,sizey
    mov ebx,sizex
    mul ebx
    mov ColArraySize,eax    ; Num 4-byte chunks

    xor eax,eax
    mov eax,Increment
    shl eax,8           ;00 00 In 00
                        ;A  R  G  B
    mov aL,Increment
    shl eax,8           ;00 In In 00
                        ;A  R  G  B
    mov aL,Increment    ;00 In In In
                        ;A  R  G  B
    
    mov lo32,eax
    mov hi32,eax
    movq mm0,hi32
    
    mov ecx,ColArraySize
    shr ecx,1           ; Num 8 byte chunks
    mov ebx,8
LD:
    movq mm1,[edi]
    psubusb mm1,mm0
    movq [edi],mm1
    add edi,ebx
    dec ecx
    jnz LD
    emms

RET

;================================================
ColArrayAddr2:    ; In: LONG ARRAY, kx,ky,sizex,sizey  Out: edi->addr
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
;===========================

