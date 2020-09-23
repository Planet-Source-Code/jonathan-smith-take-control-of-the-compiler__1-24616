	TITLE	C:\FastString\FastString.cpp
	.386P
include listing.inc
if @Version gt 510
.model FLAT
else
_TEXT	SEGMENT PARA USE32 PUBLIC 'CODE'
_TEXT	ENDS
_DATA	SEGMENT DWORD USE32 PUBLIC 'DATA'
_DATA	ENDS
CONST	SEGMENT DWORD USE32 PUBLIC 'CONST'
CONST	ENDS
_BSS	SEGMENT DWORD USE32 PUBLIC 'BSS'
_BSS	ENDS
$$SYMBOLS	SEGMENT BYTE USE32 'DEBSYM'
$$SYMBOLS	ENDS
$$TYPES	SEGMENT BYTE USE32 'DEBTYP'
$$TYPES	ENDS
_TLS	SEGMENT DWORD USE32 PUBLIC 'TLS'
_TLS	ENDS
;	COMDAT ?FastBStrReverse@@YGXPAPAG@Z
_TEXT	SEGMENT PARA USE32 PUBLIC 'CODE'
_TEXT	ENDS
FLAT	GROUP _DATA, CONST, _BSS
	ASSUME	CS: FLAT, DS: FLAT, SS: FLAT
endif
PUBLIC	?FastBStringReverse@modFastString_stub@@AAGXXZ			; FastBStrReverse
;	COMDAT ?FastBStrReverse@@YGXPAPAG@Z
_TEXT	SEGMENT
_BasicString$ = 8
_posLeft$ = -4
_posRight$ = -8
_Swap$ = -12
?FastBStringReverse@modFastString_stub@@AAGXXZ PROC NEAR			; FastBStrReverse, COMDAT

; 5    : {

	push	ebp
	mov	ebp, esp
	sub	esp, 76					; 0000004cH
	push	ebx
	push	esi
	push	edi
	lea	edi, DWORD PTR [ebp-76]
	mov	ecx, 19					; 00000013H
	mov	eax, -858993460				; ccccccccH
	rep stosd

; 6    : 	unsigned long posLeft, posRight;
; 7    : 	unsigned short Swap;
; 8    : 
; 9    : 	for (posLeft = 0, posRight = (*BasicString)[-2]/2 - 1; posLeft < posRight; posLeft++, posRight--) {

	mov	DWORD PTR _posLeft$[ebp], 0
	mov	eax, DWORD PTR _BasicString$[ebp]
	mov	ecx, DWORD PTR [eax]
	xor	eax, eax
	mov	ax, WORD PTR [ecx-4]
	cdq
	sub	eax, edx
	sar	eax, 1
	sub	eax, 1
	mov	DWORD PTR _posRight$[ebp], eax
	jmp	SHORT $L222
$L223:
	mov	edx, DWORD PTR _posLeft$[ebp]
	add	edx, 1
	mov	DWORD PTR _posLeft$[ebp], edx
	mov	eax, DWORD PTR _posRight$[ebp]
	sub	eax, 1
	mov	DWORD PTR _posRight$[ebp], eax
$L222:
	mov	ecx, DWORD PTR _posLeft$[ebp]
	cmp	ecx, DWORD PTR _posRight$[ebp]
	jae	SHORT $L224

; 10   : 		Swap = (*BasicString)[posLeft];

	mov	edx, DWORD PTR _BasicString$[ebp]
	mov	eax, DWORD PTR [edx]
	mov	ecx, DWORD PTR _posLeft$[ebp]
	mov	dx, WORD PTR [eax+ecx*2]
	mov	WORD PTR _Swap$[ebp], dx

; 11   : 		(*BasicString)[posLeft] = (*BasicString)[posRight];

	mov	eax, DWORD PTR _BasicString$[ebp]
	mov	ecx, DWORD PTR [eax]
	mov	edx, DWORD PTR _BasicString$[ebp]
	mov	eax, DWORD PTR [edx]
	mov	edx, DWORD PTR _posLeft$[ebp]
	mov	esi, DWORD PTR _posRight$[ebp]
	mov	cx, WORD PTR [ecx+esi*2]
	mov	WORD PTR [eax+edx*2], cx

; 12   : 		(*BasicString)[posRight] = Swap;

	mov	edx, DWORD PTR _BasicString$[ebp]
	mov	eax, DWORD PTR [edx]
	mov	ecx, DWORD PTR _posRight$[ebp]
	mov	dx, WORD PTR _Swap$[ebp]
	mov	WORD PTR [eax+ecx*2], dx

; 13   : 
; 14   : 	}

	jmp	SHORT $L223
$L224:

; 15   : }

	pop	edi
	pop	esi
	pop	ebx
	mov	esp, ebp
	pop	ebp
	ret	4
?FastBStringReverse@modFastString_stub@@AAGXXZ ENDP			; FastBStrReverse
_TEXT	ENDS
END
