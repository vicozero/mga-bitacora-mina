package com.mga.bitacoramina;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        registerPlugin(MgaSpeechPlugin.class);
        registerPlugin(MgaBiometricPlugin.class);
        super.onCreate(savedInstanceState);
    }
}
