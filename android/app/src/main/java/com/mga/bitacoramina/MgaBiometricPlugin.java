package com.mga.bitacoramina;

import androidx.annotation.NonNull;
import androidx.biometric.BiometricManager;
import androidx.biometric.BiometricPrompt;
import androidx.core.content.ContextCompat;
import androidx.fragment.app.FragmentActivity;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import java.util.concurrent.Executor;

@CapacitorPlugin(name = "MgaBiometric")
public class MgaBiometricPlugin extends Plugin {
    private PluginCall activeCall;

    @PluginMethod
    public void isAvailable(PluginCall call) {
        int result = BiometricManager.from(getContext()).canAuthenticate(BiometricManager.Authenticators.BIOMETRIC_WEAK | BiometricManager.Authenticators.DEVICE_CREDENTIAL);
        JSObject response = new JSObject();
        response.put("available", result == BiometricManager.BIOMETRIC_SUCCESS);
        response.put("reason", reason(result));
        call.resolve(response);
    }

    @PluginMethod
    public void authenticate(PluginCall call) {
        int result = BiometricManager.from(getContext()).canAuthenticate(BiometricManager.Authenticators.BIOMETRIC_WEAK | BiometricManager.Authenticators.DEVICE_CREDENTIAL);
        if (result != BiometricManager.BIOMETRIC_SUCCESS) {
            call.reject(reason(result));
            return;
        }

        activeCall = call;
        String title = call.getString("title", "MGA Bitacora Mina");
        String subtitle = call.getString("subtitle", "Confirma tu identidad");
        Executor executor = ContextCompat.getMainExecutor(getContext());
        BiometricPrompt prompt = new BiometricPrompt((FragmentActivity) getActivity(), executor, new BiometricPrompt.AuthenticationCallback() {
            @Override
            public void onAuthenticationError(int errorCode, @NonNull CharSequence errString) {
                rejectActiveCall(errString.toString());
            }

            @Override
            public void onAuthenticationSucceeded(@NonNull BiometricPrompt.AuthenticationResult result) {
                JSObject response = new JSObject();
                response.put("verified", true);
                resolveActiveCall(response);
            }

            @Override
            public void onAuthenticationFailed() {
                // Keep the prompt open so Android can retry.
            }
        });

        BiometricPrompt.PromptInfo info = new BiometricPrompt.PromptInfo.Builder()
            .setTitle(title)
            .setSubtitle(subtitle)
            .setAllowedAuthenticators(BiometricManager.Authenticators.BIOMETRIC_WEAK | BiometricManager.Authenticators.DEVICE_CREDENTIAL)
            .build();
        prompt.authenticate(info);
    }

    private void resolveActiveCall(JSObject response) {
        PluginCall call = activeCall;
        activeCall = null;
        if (call != null) call.resolve(response);
    }

    private void rejectActiveCall(String message) {
        PluginCall call = activeCall;
        activeCall = null;
        if (call != null) call.reject(message);
    }

    private String reason(int code) {
        switch (code) {
            case BiometricManager.BIOMETRIC_SUCCESS:
                return "";
            case BiometricManager.BIOMETRIC_ERROR_NONE_ENROLLED:
                return "No hay huella o bloqueo de pantalla registrado.";
            case BiometricManager.BIOMETRIC_ERROR_NO_HARDWARE:
                return "Este telefono no tiene sensor biometrico.";
            case BiometricManager.BIOMETRIC_ERROR_HW_UNAVAILABLE:
                return "Sensor biometrico no disponible.";
            case BiometricManager.BIOMETRIC_ERROR_SECURITY_UPDATE_REQUIRED:
                return "El telefono requiere actualizacion de seguridad para huella.";
            case BiometricManager.BIOMETRIC_ERROR_UNSUPPORTED:
                return "Huella no soportada en este telefono.";
            case BiometricManager.BIOMETRIC_STATUS_UNKNOWN:
            default:
                return "Huella no disponible.";
        }
    }
}
