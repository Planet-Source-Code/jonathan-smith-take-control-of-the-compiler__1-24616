
extern void FastBStrReverse ( unsigned short **BasicString );

void FastBStrReverse ( unsigned short **BasicString )
{
	unsigned long posLeft, posRight;
	unsigned short Swap;

	for (posLeft = 0, posRight = (*BasicString)[-2]/2 - 1; posLeft < posRight; posLeft++, posRight--) {
		Swap = (*BasicString)[posLeft];
		(*BasicString)[posLeft] = (*BasicString)[posRight];
		(*BasicString)[posRight] = Swap;

	}
}