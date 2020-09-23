; Scroll.asm   For NASM  by Robert Rayment 26/5/02

; VB
;Public Type ScrollStruc
;   sizex As Long           ' 128 - 512
;   sizey As Long           ' 128 - 384
;   PtrColArray As Long
;   PtrLineStore As Long
;   HScrollValue As Long    ' -4 to +4
;   VScrollValue As Long    ' -4 to +4
;End Type
;Public ScrollMCODE As ScrollStruc
;Public ScrollArray() As Byte  ' Array to hold machine code
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
%define PtrLineStore    [ebp-16]
%define HScrollValue    [ebp-20]
%define VScrollValue    [ebp-24]

%Define kx              [ebp-28]
%Define ky              [ebp-32]

[BITS 32]
push ebp
mov ebp, esp
sub esp,32
push esi
push edi
push ebx
;-----------
    ; Copy structure
    mov ebx,[ebp+8]
    
    movab sizex,            [ebx]
    movab sizey,            [ebx+4]
    movab PtrColArray,      [ebx+8]
    movab PtrLineStore,     [ebx+12]
    movab HScrollValue,     [ebx+16]
    movab VScrollValue,     [ebx+20]

    mov eax,HScrollValue
    cmp eax,0
    je UpDown
    jg Right

Left:
    fild dword HScrollValue
    fabs
    fistp dword HScrollValue

    mov ecx,HScrollValue
LeftAgain:
    push ecx
    Call MMXLeftScroller
    pop ecx
    dec ecx
    jnz LeftAgain

    jmp UpDown

Right:

    mov ecx,HScrollValue
RightAgain:
    push ecx
    Call MMXRightScroller
    pop ecx
    dec ecx
    jnz RightAgain

UpDown:
    mov eax,VScrollValue
    cmp eax,0
    je GETOUT
    jg Down

Up: ;-ve VScrollValue
    fild dword VScrollValue
    fabs
    fistp dword VScrollValue

    mov ecx,VScrollValue
UpAgain:
    push ecx
    Call UpScroller
    pop ecx
    dec ecx
    jnz UpAgain

    jmp GETOUT

Down:   ;+ve VScrollValue

    mov ecx,VScrollValue
DownAgain:
    push ecx
    Call DownScroller
    pop ecx
    dec ecx
    jnz DownAgain

;-----------
GETOUT:
pop ebx
pop edi
pop esi
mov esp, ebp
pop ebp
ret 16
;############################################################

;   MMX LEFT SCROLLER  HScrollValue times

MMXLeftScroller:    ;-ve HScrollValues

    ; Store left column in LineStore
    mov esi,PtrColArray
    mov edi,PtrLineStore

    mov edx,sizex
    shl edx,2           ; ColArray Vert address step

    mov ecx,sizey
LS1:
    push esi
    movsd               ; [esi]->[edi] esi+4, edi+4
    pop esi
    add esi,edx
    dec ecx
    jnz LS1

    ; Move columns to left  MMX 8 bytes 2 columns at a time
    ; Columns  1,2 <- 2,3 TO Columns sizex-3,sizex-2 <- sizex-2,sizex-1

    mov eax,2
    mov kx,eax
    mov eax,1
    mov ky,eax
    Call ColArrayAddr   ; edi -> ColArray(2,1)
    push edi
    pop esi             ; esi -> ColArray(2,1)
    mov edi,PtrColArray ; edi -> ColArray(1,1)
    
    mov edx,sizex
    shl edx,2           ; Vert address step

    mov ebx,8           ; For incr edi & esi by 8 bytes

    mov ecx,sizey
LS2:
    push ecx
    
    mov ecx,sizex
    dec ecx
    dec ecx
    shr ecx,1           ; Num 8-byte chunks

    push edi
    push esi
LS3:
    movq mm0,[esi]      ;pick up 8 bytes
    movq [edi],mm0      ;place 8 bytes 1 byte to left
    add edi,ebx
    add esi,ebx
    dec ecx
    jnz LS3
    ; Move last column
    movq mm0,[esi]      ;pick up 8 bytes
    movq [edi],mm0      ;place 8 bytes 1 byte to left

    pop esi
    pop edi
    pop ecx
    add edi,edx
    add esi,edx
    dec ecx
    jnz LS2
    
    ; Move Linestore into right column

    mov eax,sizex
    mov kx,eax
    mov eax,1
    mov ky,eax
    Call ColArrayAddr   ; edi->ColArray(sizex,1)
    mov esi,PtrLineStore

    mov edx,sizex
    shl edx,2           ; ColArray Vert address step

    mov ecx,sizey
LS4:
    push edi
    movsd               ; [esi]->[edi] esi+4, edi+4
    pop edi
    add edi,edx
    dec ecx
    jnz LS4

    emms            ;Clear FP/MMX stack
RET

;============================================================

;   MMX RIGHT SCROLLER    HScrollValue times

MMXRightScroller:
    ; Store right column in LineStore

    mov eax,sizex
    mov kx,eax
    mov eax,1
    mov ky,eax
    Call ColArrayAddr   ; edi->ColArray(sizex,1)
    push edi
    pop esi             ; esi->ColArray(sizex,1)
    mov edi,PtrLineStore

    mov edx,sizex
    shl edx,2           ; ColArray Vert address step

    mov ecx,sizey
RS1:
    push esi
    movsd               ; [esi]->[edi] esi+4, edi+4
    pop esi
    add esi,edx
    dec ecx
    jnz RS1

    ; Move columns to right
    ; Column  sizex-2,sizex-1 -> sizex -1,sizex
    ;         2,3 to 3,4
    mov eax,sizex
    mov kx,eax
    mov eax,1
    mov ky,eax
    Call ColArrayAddr   ; edi->ColArray(sizex,1)
    mov eax,4
    sub edi,eax         ; edi->ColArray(sizex-1,1)
    push edi
    pop esi
    mov eax,4
    sub esi,eax         ; esi->ColArray(sizex-2,1)

    mov edx,sizex
    shl edx,2           ; Vert address step

    mov ecx,sizey
RS2:
    push ecx
    
    mov ecx,sizex
    dec ecx
    shr ecx,1           ; Num 8-byte chunks
    mov ebx,8
    push edi
    push esi
RS3:
    movq mm0,[esi]      ;pick up 8 bytes
    movq [edi],mm0      ;place 8 bytes 1 byte to left
    sub edi,ebx
    sub esi,ebx
    dec ecx
    jnz RS3
    ; Move first column
    movq mm0,[esi]      ;pick up 8 bytes
    movq [edi],mm0      ;place 8 bytes 1 byte to left

    pop esi
    pop edi
    pop ecx
    add edi,edx
    add esi,edx
    dec ecx
    jnz RS2

    ; Move Linestore into left column

    mov esi,PtrLineStore
    mov edi,PtrColArray

    mov edx,sizex
    shl edx,2           ; ColArray Vert address step

    mov ecx,sizey
RS4:
    push edi
    movsd               ; [esi]->[edi] esi+4, edi+4
    pop edi
    add edi,edx
    dec ecx
    jnz RS4

    emms            ;Clear FP/MMX stack
RET

;============================================================

;   UP SCROLLER   VScrollValue times

UpScroller: ;-ve VScrollValue input

    ; Move top line to Linestore

    mov eax,sizey
    mov ky,eax
    mov eax,1
    mov kx,eax
    Call ColArrayAddr   ; edi->ColArray(1,sizey)
    push edi
    pop esi             ; esi->ColArray(1,sizey)
    mov edi,PtrLineStore
    mov ecx,sizex       ; Number of dwords to move
    rep movsd

    ; Move rows up
    ; ColArray rows sizey-1 -> sizey TO 1 -> 2

    mov eax,sizey
    mov ky,eax
    mov eax,sizex
    mov kx,eax
    Call ColArrayAddr   ; edi->ColArray(sizex,sizey)
    push edi
    pop esi             ; esi->ColArray(sizex,sizey)
    mov eax,sizex
    shl eax,2
    sub esi,eax         ; esi->ColArray(sizex,sizey-1)

    mov eax,sizex
    mov ebx,sizey
    dec ebx
    mul ebx
    mov ecx,eax         ; Number of dwords to move up
    std                 ; decr
    rep movsd           ; [esi]->[edi] esi-4, edi-4
    cld

    ; Move Linestore to bottom line

    mov edi,PtrColArray
    mov esi,PtrLineStore
    mov ecx,sizex       ; Number of dwords to move
    rep movsd

RET

;============================================================

;   DOWN SCROLLER     VScrollValue times

DownScroller:   ;+ve VScrollValue input

    ; Move bottom line to Linestore

    mov esi,PtrColArray
    mov edi,PtrLineStore
    mov ecx,sizex           ; Number of dwords to move      
    rep movsd

    ; Move rows down
    ; ColArray rows 2 -> 1 TO sizey -> sizey-1 

    mov edi,PtrColArray     ; edi->ColArray(1,1)
    push edi
    pop esi                 ; esi->ColArray(1,1)
    mov eax,sizex
    shl eax,2
    add esi,eax             ; esi->ColArray(1,2)

    mov eax,sizex
    mov ebx,sizey
    dec ebx
    mul ebx
    mov ecx,eax             ; Num of dwords to move down
    rep movsd               ; [esi]->[edi] esi+4, edi+4

    ; Move Linestore to top line

    mov eax,sizey
    mov ky,eax
    mov eax,1
    mov kx,eax
    Call ColArrayAddr       ; edi->ColArray(1,sizey)
    mov esi,PtrLineStore    ; esi->Linestore(1)
    mov ecx,sizex           ; Number of dwords to move 
    rep movsd

RET
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