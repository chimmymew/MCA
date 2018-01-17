


int adc_read;
long mca[1024];
char data[2];


void setup() {
  // put your setup code here, to run once:
 adc_init(ADC, SystemCoreClock, ADC_FREQ_MAX, ADC_STARTUP_FAST); 


  analogReadResolution(12); //
  //analogReference(INTERNAL2V56);
 Serial.begin(115200);
}

void loop() {
  // put your main code here, to run repeatedly:
byte val;
  if (Serial.available()) {
    val = Serial.read();
    switch (val) {
  
        case ('C'): 

         //Serial.println("Clearing array...");
         
         for (int i=0; i < 1024; i++){
         
         mca[i]=0;
         
         
         }  

         //Serial.println("Integrating...");    
         
         for (long i=0; i < 30000; i++){
         adc_read = analogRead(0)/4;
         mca[adc_read]++;
   
         } 

          //Serial.println("Data...");
          Serial.print("**");
         for (int i=0; i < 1024; i++){
         data[0] =  (mca[i] & 0xFF);
         data[1] = ((mca[i] >> 8) & 0xFF);
         
         Serial.print(data[0]);
         Serial.print(data[1]);
         } 

         Serial.println("%%");
        
        break;


        case ('I'): 

         Serial.println("Clearing array...");
         
         for (int i=0; i < 1024; i++){
         
         mca[i]=0;
         
         
         }  

         Serial.println("Integrating...");    
         
         for (long i=0; i < 30000; i++){
         adc_read = analogRead(0)/4;
         mca[adc_read]++;
   
         } 

          Serial.println("Data...");
          Serial.print("*,");
         for (int i=0; i < 1024; i++){

         Serial.print(mca[i]);
         Serial.print(",");
         } 

         Serial.println("#");
        
        break;



        case ('O'): 

         Serial.println("Clearing array...");
         
         for (int i=0; i < 1024; i++){
         
         mca[i]=0;
         
         
         }  

         Serial.println("Collecting...");    
         
         for (long i=0; i < 1024; i++){
         adc_read = analogRead(0);
         mca[i] = adc_read;
   
         } 

          Serial.println("Data...");
          Serial.print("*,");
         for (int i=0; i < 1024; i++){

         Serial.print(mca[i]);
         Serial.print(",");
         } 

         Serial.println("#");
        
        break;

        

        case ('T'): 

         Serial.println("Clearing array...");
         
         for (int i=0; i < 1024; i++){
         
         mca[i]=0;
         
         
         }  

         Serial.println("Integrating...");    
         
         for (long i=0; i < 30000; i++){
         adc_read = analogRead(0)/4;
         mca[adc_read]++;
   
         } 

          Serial.println("Statistical Table");
          Serial.println("------------");
          Serial.println("CH,Value");
         for (int i=0; i < 1024; i++){

         Serial.print(i);
         Serial.print(",");
         Serial.println(mca[i]);
         } 

         
        
        break;


        
    }
    
  }



}
