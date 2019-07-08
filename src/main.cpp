#include <Arduino.h>
#include <TheThingsNetwork.h>
#include "DHT.h"
#include <CayenneLPP.h>

// Set your AppEUI and AppKey
const char *appEui = "70B3D57ED001FABC";
const char *appKey = "5E95A080B6561090BA0463389BABCDEF";

#define loraSerial Serial1
#define debugSerial Serial

// Replace REPLACE_ME with TTN_FP_EU868 or TTN_FP_US915
#define freqPlan TTN_FP_EU868

TheThingsNetwork ttn(loraSerial, debugSerial, freqPlan);
DHT dht;
CayenneLPP lpp(51);

void setup()
{
  loraSerial.begin(57600);
  debugSerial.begin(9600);

  // Wait a maximum of 10s for Serial Monitor
  while (!debugSerial && millis() < 10000)
    ;

  dht.setup(A0);

  debugSerial.println("-- STATUS");
  ttn.showStatus();

  debugSerial.println("-- JOIN");
  ttn.join(appEui, appKey);
  delay(2000);
}

void loop()
{
  debugSerial.println("-- LOOP");

  float h = dht.getHumidity();
  float t = dht.getTemperature();

  if (isnan(h) || isnan(t))
  {
    Serial.println(F("Failed to read from DHT sensor!"));
    return;
  }

  Serial.print(F("Humidity: "));
  Serial.print(h);
  Serial.print(F("%  Temperature: "));
  Serial.print(t);
  Serial.println(F("°C "));

  lpp.reset();
  lpp.addTemperature(1, t);
  lpp.addRelativeHumidity(2, h );

  ttn.sendBytes(lpp.getBuffer(), lpp.getSize());

  delay(5000);
}