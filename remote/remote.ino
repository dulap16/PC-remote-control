// REMOTE CONTROLLER - ARDUINO HALF

// CH-        69 
// CH         70 
// CH+        71
// PREV       68 
// NEXT       64 
// PLAY/PAUSE 67
// VOL-       7 
// VOL+       21 
// EQ         9
// 0          22 
// 100+       25 
// 200+       13
// 1          12 
// 2          24 
// 3          94
// 4          8 
// 5          28 
// 6          90
// 7          66 
// 8          82 
// 9          74

#include <IRremote.h>

#define IR_RECEIVE_PIN 7

void setup(){
  Serial.begin(9600);
  IrReceiver.begin(IR_RECEIVE_PIN);
}

void loop(){
  if (IrReceiver.decode()){
      IrReceiver.resume();
      Serial.println(IrReceiver.decodedIRData.command);
  }
}