	TITLE	E:\Compile Controller\Online Package\InLine Assembly\Math.bas
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
text$1	SEGMENT PARA USE32 PUBLIC ''
text$1	ENDS
;	COMDAT ?DivideBy2Normally@Math@@AAGXXZ
text$1	SEGMENT PARA USE32 PUBLIC ''
text$1	ENDS
;	COMDAT ?DivideBy2ByShifting@Math@@AAGXXZ
text$1	SEGMENT PARA USE32 PUBLIC ''
text$1	ENDS
FLAT	GROUP _DATA, CONST, _BSS
	ASSUME	CS: FLAT, DS: FLAT, SS: FLAT
endif
PUBLIC	?DivideBy2Normally@Math@@AAGXXZ			; Math::DivideBy2Normally
PUBLIC	__real@8@40008000000000000000
EXTRN	@__vbaFpI4:NEAR
EXTRN	___vbaErrorOverflow:NEAR
EXTRN	___vbaChkstk:NEAR
EXTRN	__fltused:NEAR
EXTRN	__adjust_fdiv:DWORD
EXTRN	__adj_fdiv_m64:NEAR
EXTRN	___vbaFPException:NEAR
;	COMDAT __real@8@40008000000000000000
; File E:\Compile Controller\Online Package\InLine Assembly\Math.bas
CONST	SEGMENT
__real@8@40008000000000000000 DQ 04000000000000000r ; 2
CONST	ENDS
;	COMDAT ?DivideBy2Normally@Math@@AAGXXZ
text$1	SEGMENT
_lngIterations$ = 12
_lngDividend$ = 8
_xIterations$ = -4
_DivideBy2Normally$ = -8
_unnamed_var1$ = -12
_unnamed_var2$ = -16					; ** YOU MUST BREAK OUT THE "UNNAMED VARS" **
?DivideBy2Normally@Math@@AAGXXZ PROC NEAR		; Math::DivideBy2Normally, COMDAT

; 3    : Public Function DivideBy2Normally(lngDividend As Long, lngIterations As Long) As Long

	push	ebp
	mov	ebp, esp
	push	24					; 00000018H
	pop	eax
	call	___vbaChkstk

; 4    :     Dim xIterations As Long
; 5    :     For xIterations = 1 To lngIterations

	mov	eax, DWORD PTR _lngIterations$[ebp]
	mov	eax, DWORD PTR [eax]
	mov	DWORD PTR _unnamed_var1$[ebp], eax
	mov	DWORD PTR _unnamed_var2$[ebp], 1		; ** unnamed 2 is loop increment **

	mov	DWORD PTR _xIterations$[ebp], 1
	jmp	SHORT $L28
$L27:

; 6    :         DivideBy2Normally = lngDividend / 2

	mov	eax, DWORD PTR _xIterations$[ebp]
	add	eax, DWORD PTR _unnamed_var2$[ebp]		; ** unnamed 2 is loop increment **
	jo	SHORT $L19
	mov	DWORD PTR _xIterations$[ebp], eax
$L28:
	mov	eax, DWORD PTR _xIterations$[ebp]
	cmp	eax, DWORD PTR _unnamed_var1$[ebp]
	jg	SHORT $L18
	mov	eax, DWORD PTR _lngDividend$[ebp]
	fild	DWORD PTR [eax]
	fstp	QWORD PTR -24+[ebp]
	fld	QWORD PTR -24+[ebp]
	cmp	DWORD PTR __adjust_fdiv, 0
	jne	SHORT $L54
	fdiv	QWORD PTR __real@8@40008000000000000000
	jmp	SHORT $L55
$L54:
	push	DWORD PTR __real@8@40008000000000000000+4
	push	DWORD PTR __real@8@40008000000000000000
	call	__adj_fdiv_m64
$L55:
	fnstsw	ax
	test	al, 13					; 0000000dH
	jne	SHORT $L49
	call	@__vbaFpI4
	mov	DWORD PTR _DivideBy2Normally$[ebp], eax

; 7    :     Next

	jmp	SHORT $L27
$L18:

; 8    : End Function

	mov	eax, DWORD PTR _DivideBy2Normally$[ebp]
	leave
	ret	8
$L49:
	jmp	___vbaFPException
$L19:
	call	___vbaErrorOverflow
?DivideBy2Normally@Math@@AAGXXZ ENDP			; Math::DivideBy2Normally
text$1	ENDS
PUBLIC	?DivideBy2ByShifting@Math@@AAGXXZ		; Math::DivideBy2ByShifting
;	COMDAT ?DivideBy2ByShifting@Math@@AAGXXZ
text$1	SEGMENT
_lngIterations$ = 12
_lngDividend$ = 8
_DivideBy2ByShifting$ = -4
_xIterations$ = -8
_unnamed_var1$ = -12
_unnamed_var2$ = -16						; ** BREAK OUT SECOND UNNAMED VAR **
?DivideBy2ByShifting@Math@@AAGXXZ PROC NEAR		; Math::DivideBy2ByShifting, COMDAT

; 10   : Public Function DivideBy2ByShifting(lngDividend As Long, lngIterations As Long) As Long

	push	ebp
	mov	ebp, esp
	push	24					; 00000018H
	pop	eax
	call	___vbaChkstk

; 11   :     Dim xIterations As Long
; 12   :     For xIterations = 1 To lngIterations

	mov	eax, DWORD PTR _lngIterations$[ebp]
	mov	eax, DWORD PTR [eax]
	mov	DWORD PTR _unnamed_var1$[ebp], eax
	mov	DWORD PTR _unnamed_var2$[ebp], 1		; ** unnamed 2 is loop increment **
	mov	DWORD PTR _xIterations$[ebp], 1
	jmp	SHORT $L42
$L41:

; 13   :         DivideBy2ByShifting = lngDividend / 2

	mov	eax, DWORD PTR _xIterations$[ebp]
	add	eax, DWORD PTR _unnamed_var2$[ebp]		; ** unnamed 2 is loop increment **
	jo	SHORT $L33
	mov	DWORD PTR _xIterations$[ebp], eax
$L42:
	mov	eax, DWORD PTR _xIterations$[ebp]
	cmp	eax, DWORD PTR _unnamed_var1$[ebp]
	jg	SHORT $L32
	mov	eax, DWORD PTR _lngDividend$[ebp]

; *** THIS IS THE OLD ROUNDING CODE WE ARE NIXING ***
;
;	fild	DWORD PTR [eax]
;	fstp	QWORD PTR -24+[ebp]
;	fld	QWORD PTR -24+[ebp]
;	cmp	DWORD PTR __adjust_fdiv, 0
;	jne	SHORT $L62
;	fdiv	QWORD PTR __real@8@40008000000000000000
;	jmp	SHORT $L63
;$L62:
;	push	DWORD PTR __real@8@40008000000000000000+4
;	push	DWORD PTR __real@8@40008000000000000000
;	call	__adj_fdiv_m64
;$L63:
;	fnstsw	ax
;	test	al, 13					; 0000000dH
;	jne	SHORT $L61
;	call	@__vbaFpI4

; *** THESE TWO INSTRUCTIONS NOW DO THE DIVISION ***
	mov	eax, DWORD PTR [eax]
	sar	eax, 1

	mov	DWORD PTR _DivideBy2ByShifting$[ebp], eax

; 14   :     Next

	jmp	SHORT $L41
$L32:

; 15   : End Function

	mov	eax, DWORD PTR _DivideBy2ByShifting$[ebp]
	leave
	ret	8
$L61:
	jmp	___vbaFPException
$L33:
	call	___vbaErrorOverflow
?DivideBy2ByShifting@Math@@AAGXXZ ENDP			; Math::DivideBy2ByShifting
text$1	ENDS
END
