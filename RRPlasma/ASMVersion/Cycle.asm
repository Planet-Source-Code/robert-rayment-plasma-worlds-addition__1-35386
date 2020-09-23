; Cycle.asm   For NASM by Robert Rayment 26/5/02

; VB
;'Color cycling MCode Structure
;Public Type CycleStruc
;   sizex As Long           ' 128 - 512
;   sizey As Long           ' 128 - 384
;   PtrColArray As Long
;   CycleValue As Long    ' -4 to +4
;End Type
;Public CycleMCODE As EffectsStruc
;Public CycleArray() As Byte  ' Array to hold machine code
;'res = CallWindowProc(ptrMC, ptrStruc, 0, 0&, 0&)
;                              +8       +12 +16 +20

%macro movab 2      ; name & num of parameters
  push dword %2     ; 2nd param
  pop dword %1      ; 1st param
%endmacro           ; use  movab %1,%2
; Allows eg movab bmW,[ebx+4]


%define sizex           [ebp-4]
%define sizey           [ebp-8]
%define PtrColArray     [ebp-12]
%define CycleValue      [ebp-16]

%Define ABSCycleValue   [ebp-20]

; For MMX
%Define lo32 [ebp-24]
%Define hi32 [ebp-28]

[BITS 32]
push ebp
mov ebp, esp
sub esp,28
push esi
push edi
push ebx
;-----------
    ; Copy structure
    mov ebx,[ebp+8]
    
    movab sizex,            [ebx]
    movab sizey,            [ebx+4]
    movab PtrColArray,      [ebx+8]
    movab CycleValue,       [ebx+12]

    mov eax, CycleValue
    cmp eax,0
    jz near GETOUT

CYCLE:

    mov eax,sizey
    mov ebx,sizex
    mul ebx
    shr eax,1           ; sixex*sizey/2  8-byte chunks
    mov ecx,eax         ; ecx = number of 8-byte chunks
    
    mov edi,PtrColArray        ; edi-> PtrColArray(1,1)
    
    mov eax,CycleValue  ; +/- 1,2,3 or 4
    fild dword CycleValue
    fabs
    fistp dword ABSCycleValue
    mov eax,ABSCycleValue
    
    cmp eax,1
    jne CSp2
    mov eax,01010101h   ;Make 8 bytes
    jmp LoadSpeed
CSp2:
    cmp eax,2
    jne CSp3
    mov eax,02020202h   ;Make 8 bytes       ; double speed
    jmp LoadSpeed
CSp3:
    cmp eax,3
    jne CSp4
    mov eax,03030303h   ;Make 8 bytes       ; triple speed

CSp4:
    mov eax,04040404h   ;Make 8 bytes       ; quad speed

LoadSpeed:
    mov lo32,eax        ;to add/subtract 1,2,3 or 4 from each byte
    mov hi32,eax
    
    movq mm1,hi32

    mov ebx,8

NextChunk:
    
    ; pick up 8 bytes
    movq mm0,[edi]

    mov eax, CycleValue
    cmp eax,0
    jg AddColor    

    ; subtract 1,2,3 or 4 from each byte with overlap ie 0->255
    psubb mm0,mm1       ;sub 8 bytes at a time
    jmp PutBack

AddColor:
    ; add 1,2,3 or 4 from each byte with overlap ie 0->255
    paddb mm0,mm1       ;add 8 bytes at a time
    
PutBack:    

    ; put back 8 bytes
    movq [edi],mm0      ;with overlap ie 0->255

    add edi,ebx

    dec ecx
    jnz NextChunk
    emms            ;Clear FP/MMX stack

;-----------
GETOUT:
pop ebx
pop edi
pop esi
mov esp, ebp
pop ebp
ret 16
