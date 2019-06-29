/*
 * conio.h
 *
 * Low level console I/O functions. Pretty please try to use the ANSI
 * standard ones if you are writing new code.
 *
 * This file is part of the Mingw32 package.
 *
 * Contributors:
 *  Created by Colin Peters <colin@bird.fu.is.saga-u.ac.jp>
 *
 *  THIS SOFTWARE IS NOT COPYRIGHTED
 *
 *  This source code is offered for use in the public domain. You may
 *  use, modify or distribute it freely.
 *
 *  This code is distributed in the hope that it will be useful but
 *  WITHOUT ANY WARRANTY. ALL WARRANTIES, EXPRESS OR IMPLIED ARE HEREBY
 *  DISCLAIMED. This includes but is not limited to warranties of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 *
 * $Revision: 1.4 $
 * $Author: earnie $
 * $Date: 2003/02/21 21:19:51 $
 *
 */

#ifndef	__STRICT_ANSI__

#ifndef	_CONIO_H_
#define	_CONIO_H_

/* All the headers include this file. */
#include <_mingw.h>

#include <stdio.h>

#ifndef RC_INVOKED

#ifdef	__cplusplus
extern "C" {
#endif

_CRTIMP char* __cdecl	_cgets (char*);
_CRTIMP int __cdecl	_cprintf (const char*, ...);
_CRTIMP int __cdecl	_cputs (const char*);
_CRTIMP int __cdecl	_cscanf (char*, ...);

_CRTIMP int __cdecl	_getch (void);
_CRTIMP int __cdecl	_getche (void);
_CRTIMP int __cdecl	_kbhit (void);
_CRTIMP int __cdecl	_putch (int);
_CRTIMP int __cdecl	_ungetch (int);
#define _NOCURSOR      0
#define _SOLIDCURSOR   1
#define _NORMALCURSOR  2
#ifndef _WINCON_H
typedef struct _COORD { short x;short y; } COORD;
typedef struct _CONSOLE_CURSOR_INFO { int dwSize; int bVisible; } CONSOLE_CURSOR_INFO,*PCONSOLE_CURSOR_INFO;
void * _stdcall GetStdHandle(long unsigned int);
int _stdcall SetConsoleCursorPosition(void *,_COORD);
int _stdcall SetConsoleCursorInfo(void *,const _CONSOLE_CURSOR_INFO *);
#endif
void gotoxy(int x,int y) { SetConsoleCursorPosition(GetStdHandle(0xfffffff5),(COORD){x,y}); }
void _setcursortype(int i) { CONSOLE_CURSOR_INFO c={100,i}; SetConsoleCursorInfo(GetStdHandle(0xfffffff5),&c);}
void clrscr() {int i;for(i=0;i<40;i++) puts("");gotoxy(0,0);}

#ifndef	_NO_OLDNAMES

_CRTIMP int __cdecl	getch (void);
_CRTIMP int __cdecl	getche (void);
_CRTIMP int __cdecl	kbhit (void);
_CRTIMP int __cdecl	putch (int);
_CRTIMP int __cdecl	ungetch (int);

#endif	/* Not _NO_OLDNAMES */


#ifdef	__cplusplus
}
#endif

#endif	/* Not RC_INVOKED */

#endif	/* Not _CONIO_H_ */

#endif	/* Not __STRICT_ANSI__ */
