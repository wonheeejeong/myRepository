#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <time.h>
#define K 1000
#define TRY 100000

typedef struct _Item{

    long long price;
    int eqLv;
    int lv_SF;
    long long totalCost;
    int totalDestroyed;

}Item;

static double probSuccessList[25] =
{
    0.95, 0.90, 0.85, 0.85, 0.80,
    0.75, 0.70, 0.65, 0.60, 0.55,
    0.50, 0.45, 0.40, 0.35, 0.30,
    0.30, 0.30, 0.30, 0.30, 0.30,
    0.30, 0.30, 0.03, 0.02, 0.01
};

static double proDestroyedList[25] =
{
    0.00,   0.00,   0.00,   0.00,   0.00,
    0.00,   0.00,   0.00,   0.00,   0.00,
    0.00,   0.00,   0.006,  0.013,  0.014,
    0.021,  0.021,  0.021,  0.028,  0.028,
    0.07,   0.07,   0.194,  0.294,  0.396
};

int failStack = 0;
char buf[50];


char *commify(double val, char *buf, int round) {
  static char *result;
  char *nmr;
  int dp, sign;


  result = buf;

  if (round < 0)                        /*  Be sure round-off is positive  */
    round = -round;

  nmr = fcvt(val, round, &dp, &sign);   /*  Convert number to a string     */

  if (sign)                             /*  Prefix minus sign if negative  */
    *buf++ = '-';

  if (dp <= 0){                         /*  Check if number is less than 1 */
    if (dp < -round)                    /*  Set dp to max(dp, -round)      */
      dp = -round;
    *buf++ = '0';                       /*  Prefix with "0."               */
    *buf++ = '.';
    while (dp++)                        /*  Write zeros following decimal  */
      *buf++ = '0';                     /*     point                       */
  }
  else {                                /*  Number is >= 1, commify it     */
    while (dp--){
      *buf++ = *nmr++;
      if (dp % 3 == 0)
        *buf++ = dp ? ',' : '.';
    }
  }

  strcpy(buf, nmr);                     /*  Append rest of digits         */
  return result;                        /*  following dec pt              */
}


int cost_SF(int eqLv, int lv_SF){
    int cost;
    if(lv_SF <= 9)
        cost = 1000+pow(eqLv, 3)*(lv_SF+1)/25;
    else if(lv_SF >= 10 && lv_SF <= 14)
        cost = 1000+pow(eqLv, 3)*pow((lv_SF+1),2.7)/400;
    else
        cost = 1000+pow(eqLv, 3)*pow((lv_SF+1),2.7)/200;
    cost = cost - cost%100;
    return cost;

}

int getResultSF(int lv_SF, int SC, int protect){

    if(lv_SF == 25)
        return 0; //lv_SF is already 25

    int tmp = 0;

    double probSuccess = probSuccessList[lv_SF];
    if(SC) probSuccess *= 1.05;
    if(failStack == 2){
            //printf("chance time!\n");
            probSuccess = 1;
            failStack = 0;
    }

    double probDestroyed = proDestroyedList[lv_SF];
    if(protect == 1 && (lv_SF >=12 && lv_SF < 17)){
        probDestroyed = 0;
        tmp += 3;
    }
    double probFail = 1 - probSuccess - probDestroyed;

    int k = rand() % K + 1;
    int boundary1 = K - (int)(K*probSuccess)+1;
    int boundary2 = boundary1 - (int)(K*probDestroyed);


    if(k >= boundary1 && k <= K)
        tmp += 1; //success( 4 = success with protect)
    else if(k >= boundary2 && k < boundary1)
        tmp += 2; //destroyed
    else if(k < boundary2)
        tmp += 3; //fail ( 6 = fail with protect)

    return tmp;
}

void changeSF(Item *item, int result){

    if(result == 0) printf("error!\n");
    if(result == 1 || result == 4){
        //printf("Success\n");
        if(result == 4 && ( item->lv_SF >= 12 && item->lv_SF < 17) ){
            //printf("!!!protected!!!\n");
            item->totalCost += 2*cost_SF(item->eqLv, item->lv_SF);
        }
        else
            item->totalCost += cost_SF(item->eqLv, item->lv_SF);
        item->lv_SF++;
        failStack = 0;
    }
    if(result == 2){
        //printf("@@@Destroyed@@@\n");
        item->totalCost += cost_SF(item->eqLv, item->lv_SF)+item->price;
        item->totalDestroyed++;
        item->lv_SF = 12;
        failStack = 0;
    }
    if(result == 3 || result == 6){
        //printf("Failed\n");
        if(result == 6 && ( item->lv_SF >= 12 && item->lv_SF < 17) ){
            //printf("!!!protected!!!\n");
            item->totalCost += 2*cost_SF(item->eqLv, item->lv_SF);
        }
        else
            item->totalCost += cost_SF(item->eqLv, item->lv_SF);
        if( (item->lv_SF > 10) && ((item->lv_SF != 15) && (item->lv_SF != 20)) ){
            item->lv_SF--;
            failStack++;
        }
    }
    //printf("current Cost : %s\n", commify((item->totalCost), buf, 0));

}



int main(void)
{

    srand( (unsigned)time(NULL) );

    for(int p = 0; p < 2; p++){
        for(int q = 0; q < 2; q++){
            if(p == 0 && q == 0)
                printf("No starCatch, No protect\n");
            if(p == 0 && q == 1)
                printf("No starCatch, protect\n");
            if(p == 1 && q == 0)
                printf("starCatch, No protect\n");
            if(p == 1 && q == 1)
                printf("starCatch, protect\n");

            long long resultCost[TRY];
            long long resultCostSum = 0;
            int resultDestroyed[TRY];
            int resultDestroyedSum = 0;

            for(int i = 0; i < TRY; i++){

                Item sampleItem = {3550000000, 200, 0, 0, 0};

                while(sampleItem.lv_SF < 22){
                    //printf("%d성 -> %d성 스타포스 시도\n", sampleItem.lv_SF, sampleItem.lv_SF + 1);
                    changeSF(&sampleItem, getResultSF(sampleItem.lv_SF, p, q));
                    //printf("%시도 결과 : %d성\n\n", sampleItem.lv_SF);
                }
                resultCost[i] = sampleItem.totalCost;
                resultCostSum += sampleItem.totalCost;
                resultDestroyed[i] = sampleItem.totalDestroyed;
                resultDestroyedSum += sampleItem.totalDestroyed;
                //printf("총 %d회 도전 중 %d번째 시도비용 : %s메소\n"  , TRY, i+1, commify(resultCost[i], buf, 0));
                //printf("총 %d회 도전 중 %d번째 파괴횟수 : %s회\n\n", TRY, i+1, commify(resultDestroyed[i], buf, 0));

            }

            printf("\n\n평균 시도비용 : %s메소\n" ,commify(resultCostSum/TRY, buf, 0));
            printf("평균 파괴횟수 : %s회\n" ,     commify((double)(resultDestroyedSum)/TRY, buf, 3));

            long long temp;
            for(int i = 0 ; i < TRY-1 ; i ++) {
                for(int j = i+1 ; j < TRY ; j ++) {
                    if(resultCost[i] > resultCost[j]) {
                        temp = resultCost[j];
                        resultCost[j] = resultCost[i];
                        resultCost[i] = temp;
                   }
                }
            }


            for(int i = 0 ; i < TRY-1 ; i ++) {
                for(int j = i+1 ; j < TRY ; j ++) {
                    if(resultDestroyed[i] > resultDestroyed[j]) {
                        temp = resultDestroyed[j];
                        resultDestroyed[j] = resultDestroyed[i];
                        resultDestroyed[i] = temp;
                   }
                }
            }

            printf("1등(상위 0.001퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[0], buf, 0), resultDestroyed[0]);
            printf("100등(상위 0.1퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[99], buf, 0), resultDestroyed[99]);
            printf("1000등(상위 1퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[999], buf, 0), resultDestroyed[999]);
            printf("10000등(상위 10퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[9999], buf, 0), resultDestroyed[9999]);
            printf("50000등(상위 50퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[49999], buf, 0), resultDestroyed[49999]);
            printf("75000등(상위 75퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[74999], buf, 0), resultDestroyed[74999]);
            printf("90000등(상위 90퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[89999], buf, 0), resultDestroyed[89999]);
            printf("100000등(상위 100퍼센트)\n");
            printf("%s메소\n%d번\n\n", commify(resultCost[99999], buf, 0), resultDestroyed[99999]);

            }
    }


	return 0;
}
