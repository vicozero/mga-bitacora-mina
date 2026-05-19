package com.mga.bitacoramina;

import android.Manifest;
import android.content.Intent;
import android.os.Bundle;
import android.speech.RecognitionListener;
import android.speech.RecognizerIntent;
import android.speech.SpeechRecognizer;
import com.getcapacitor.JSObject;
import com.getcapacitor.PermissionState;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import com.getcapacitor.annotation.Permission;
import com.getcapacitor.annotation.PermissionCallback;
import java.util.ArrayList;
import java.util.Locale;

@CapacitorPlugin(
    name = "MgaSpeech",
    permissions = {
        @Permission(alias = "microphone", strings = { Manifest.permission.RECORD_AUDIO })
    }
)
public class MgaSpeechPlugin extends Plugin {
    private SpeechRecognizer recognizer;
    private PluginCall activeCall;

    @PluginMethod
    public void isSupported(PluginCall call) {
        JSObject result = new JSObject();
        result.put("supported", SpeechRecognizer.isRecognitionAvailable(getContext()));
        call.resolve(result);
    }

    @PluginMethod
    public void start(PluginCall call) {
        if (!SpeechRecognizer.isRecognitionAvailable(getContext())) {
            call.reject("Reconocimiento de voz no disponible en este telefono.");
            return;
        }

        if (getPermissionState("microphone") != PermissionState.GRANTED) {
            requestPermissionForAlias("microphone", call, "microphonePermsCallback");
            return;
        }

        startListening(call);
    }

    @PluginMethod
    public void stop(PluginCall call) {
        getActivity().runOnUiThread(() -> cancelActiveCall("Escucha cancelada."));
        call.resolve();
    }

    @PermissionCallback
    private void microphonePermsCallback(PluginCall call) {
        if (getPermissionState("microphone") == PermissionState.GRANTED) {
            startListening(call);
            return;
        }
        call.reject("Permiso de microfono denegado.");
    }

    private void startListening(PluginCall call) {
        getActivity().runOnUiThread(() -> {
            destroyRecognizer();
            activeCall = call;
            recognizer = SpeechRecognizer.createSpeechRecognizer(getContext());
            recognizer.setRecognitionListener(new RecognitionListener() {
                @Override
                public void onReadyForSpeech(Bundle params) {}

                @Override
                public void onBeginningOfSpeech() {}

                @Override
                public void onRmsChanged(float rmsdB) {}

                @Override
                public void onBufferReceived(byte[] buffer) {}

                @Override
                public void onEndOfSpeech() {}

                @Override
                public void onPartialResults(Bundle partialResults) {}

                @Override
                public void onEvent(int eventType, Bundle params) {}

                @Override
                public void onError(int error) {
                    rejectActiveCall(errorMessage(error));
                }

                @Override
                public void onResults(Bundle results) {
                    ArrayList<String> matches = results.getStringArrayList(SpeechRecognizer.RESULTS_RECOGNITION);
                    String transcript = matches != null && !matches.isEmpty() ? matches.get(0) : "";
                    JSObject result = new JSObject();
                    result.put("transcript", transcript);
                    resolveActiveCall(result);
                }
            });

            String language = call.getString("language", "es-MX");
            Boolean offline = call.getBoolean("offline", false);
            Intent intent = new Intent(RecognizerIntent.ACTION_RECOGNIZE_SPEECH);
            intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE_MODEL, RecognizerIntent.LANGUAGE_MODEL_FREE_FORM);
            intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE, language);
            intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE_PREFERENCE, language);
            intent.putExtra(RecognizerIntent.EXTRA_PARTIAL_RESULTS, false);
            intent.putExtra(RecognizerIntent.EXTRA_MAX_RESULTS, 1);
            intent.putExtra(RecognizerIntent.EXTRA_PREFER_OFFLINE, offline != null && offline);
            recognizer.startListening(intent);
        });
    }

    private void resolveActiveCall(JSObject result) {
        PluginCall call = activeCall;
        activeCall = null;
        destroyRecognizer();
        if (call != null) call.resolve(result);
    }

    private void rejectActiveCall(String message) {
        PluginCall call = activeCall;
        activeCall = null;
        destroyRecognizer();
        if (call != null) call.reject(message);
    }

    private void cancelActiveCall(String message) {
        if (recognizer != null) {
            recognizer.cancel();
        }
        rejectActiveCall(message);
    }

    private void destroyRecognizer() {
        if (recognizer != null) {
            recognizer.destroy();
            recognizer = null;
        }
    }

    private String errorMessage(int error) {
        switch (error) {
            case SpeechRecognizer.ERROR_AUDIO:
                return "Error de audio del microfono.";
            case SpeechRecognizer.ERROR_CLIENT:
                return "No se pudo iniciar el reconocimiento de voz.";
            case SpeechRecognizer.ERROR_INSUFFICIENT_PERMISSIONS:
                return "Permiso de microfono denegado.";
            case SpeechRecognizer.ERROR_NETWORK:
            case SpeechRecognizer.ERROR_NETWORK_TIMEOUT:
                return "Voz no disponible sin servicio de reconocimiento instalado.";
            case SpeechRecognizer.ERROR_NO_MATCH:
            case SpeechRecognizer.ERROR_SPEECH_TIMEOUT:
                return "No se detecto voz.";
            case SpeechRecognizer.ERROR_RECOGNIZER_BUSY:
                return "El reconocimiento de voz esta ocupado. Intenta de nuevo.";
            case SpeechRecognizer.ERROR_SERVER:
                return "Servicio de voz no disponible.";
            default:
                return String.format(Locale.US, "No se pudo escuchar. Codigo %d.", error);
        }
    }

    @Override
    protected void handleOnDestroy() {
        destroyRecognizer();
        super.handleOnDestroy();
    }
}
