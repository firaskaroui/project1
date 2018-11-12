#include<stdio.h>
#include "fonctions.h"
int main() {
int a;
int b;
printf("\nSaisie du premier Nombre :");
scanf("%d", a);
printf("\nsaisie du second Nombre:");
scanf("%d", b);
printf("\n%d + %d = %d", a, b, sommer(a, b));
printf("\n%d + %d = %d", a, b, soustraire(a, b));
printf("\n%d / %d = %f\n", a, b, diviser(a, b));
return (0);
}

