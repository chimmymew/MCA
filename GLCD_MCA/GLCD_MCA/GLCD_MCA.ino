
#include <openGLCD.h>

int adc_read;
int mca[128];
int TempData;
int TempData1;

gText textAreaTop(textAreaTOP);

void setup() {
  // put your setup code here, to run once:
  bitClear(ADCSRA,ADPS0); 
  bitSet(ADCSRA,ADPS1); 
  bitClear(ADCSRA,ADPS2);
  
 Serial.begin(115200);
  GLCD.Init(NON_INVERTED);

 textAreaTop.SelectFont(System5x7);
// gText textAreaBOTTOM(textAreaBOTTOM);
         textAreaTop.println(F("   Mew's Microsys  "));
         textAreaTop.println(F(" 128CH mini-MCA Scope "));
 
}

void loop() {
  // put your main code here, to run repeatedly:

 for (int i=0; i < 128; i++){
         
         mca[i]=0;
                       
         }

      for (int i=0; i < 16384; i++){
         adc_read = analogRead(6)/8;
         mca[adc_read]++;
         
         } 

         
         
 //GLCD.ClearScreen();
   //textAreaTop.println(F(" mini-MCA Scope "));
   //textAreaTop.println(F(" mini-MCA Scope "));
 for (int i=0; i < 126; i++){

         TempData = mca[i]/64;
         TempData1 = mca[i+1]/64;
         GLCD.DrawLine(i,63 - TempData,i+1 ,63 - TempData1);
         
         } 

         delay(500);
               
         //GLCD.Init(INVERTED);
         //GLCD.SetDisplayMode(INVERTED);  
  for (int i=0; i < 126; i++){

         TempData = mca[i]/64;
         TempData1 = mca[i+1]/64;
         GLCD.DrawLine(i,63 - TempData,i+1 ,63 - TempData1,WHITE);
         
         } 
         //GLCD.SetDisplayMode(NON_INVERTED);

          //GLCD.Init(NON_INVERTED);

         
       
}
