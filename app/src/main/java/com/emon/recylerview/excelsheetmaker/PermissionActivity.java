package com.emon.recylerview.excelsheetmaker;

import android.Manifest;
import android.app.Activity;
import android.content.DialogInterface;
import android.content.pm.PackageManager;
import android.support.annotation.NonNull;
import android.support.annotation.RequiresPermission;
import android.support.v4.app.ActivityCompat;
import android.support.v7.app.AlertDialog;
import android.support.v7.app.AppCompatActivity;

import java.util.ArrayList;
import java.util.List;

public abstract class PermissionActivity extends AppCompatActivity {

    public interface RequestCode {
        int PERMISSION_READ_PHONE_STATE = 0;
        int PERMISSION_FINE_LOCATION = 1;
        int PERMISSION_SEND_SMS = 2;
        int PERMISSION_MULTIPLE = 3;
        int PERMISSION_WRITE_EXTERNAL_STORAGE = 4;
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions,
                                           @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);

        // If multiple permissions were requested in one call, check if they were all granted.
        if (requestCode == RequestCode.PERMISSION_MULTIPLE) {
            boolean allPermissionsGranted = true;
            for (int grantResult : grantResults) {
                if (grantResult != PackageManager.PERMISSION_GRANTED) {
                    allPermissionsGranted = false;
                }
            }

            if (allPermissionsGranted) {
                onAllPermissionsGranted(permissions);
                return;
            }
        }

        // Else, check each one if it was granted/denied/blocked.
        for (int i = 0; i < permissions.length; i++) {
            if (grantResults[i] == PackageManager.PERMISSION_GRANTED) {
                // User granted permission.
                onPermissionGranted(permissions[i]);
            } else {
                if (ActivityCompat.shouldShowRequestPermissionRationale(this, permissions[i])) {
                    // User denied permission.
                    onPermissionDenied(permissions[i]);
                } else {
                    // User denied permission and checked 'never ask again'.
                    onPermissionBlocked(permissions[i]);
                }
            }
        }
    }

    /**
     * Checks if the app has the given permission(s).
     * <p>
     * If not, it will request them.
     * <p>
     * The method is called `checkHasPermission` to avoid the linter showing a warning in the
     * child class when it's delegating permission checks to its parent class. See
     * http://stackoverflow.com/questions/36031218/check-android-permissions-in-a
     * -method/36193309#36193309 for details.
     */
    public boolean checkHasPermission(int requestCode, String... permissions) {
        if (requestCode==RequestCode.PERMISSION_MULTIPLE&&!(permissions.length > 0)) {
            throw new IllegalArgumentException("must request at least one permission");
        }

        if (requestCode == RequestCode.PERMISSION_MULTIPLE) {
            List<String> permissions_ = new ArrayList<>();

            for (String permission : permissions) {
                if (ActivityCompat.checkSelfPermission(this, permission) !=
                        PackageManager.PERMISSION_GRANTED) {
                    permissions_.add(permission);
                }
            }

            if (!permissions_.isEmpty()) {
                requestPermissions(this, permissions_.toArray(new String[permissions_.size()]), requestCode);
                return false;
            } else {
                return true;
            }
        } else {
            if (ActivityCompat.checkSelfPermission(this, permissions[0]) !=
                    PackageManager.PERMISSION_GRANTED) {
                requestPermissions(this, permissions, requestCode);
                return false;
            } else {
                return true;
            }
        }
    }

    /**
     * Requests the given permissions.
     */
    private void requestPermissions(Activity activity, String permissions[], int resultCode) {
        showRequestPermissionsDialog(activity, permissions, resultCode);
    }

    /**
     * Called when a rationale (explanation why a permission is needed) should be shown to the user.
     * <p>
     * If the user clicks the positive button, the permission is requested again, otherwise the
     * dialog is dismissed.
     */
    public void showRationaleDialog(Activity activity, String permission, String message,
                                    int resultCode) {
        new AlertDialog.Builder(activity)
                .setMessage(message)
                .setPositiveButton("ok", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) {
                        PermissionActivity.this.showRequestPermissionDialog(activity, permission, resultCode);
                    }
                })
                .setNegativeButton("not now", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) { /* Do nothing */ }
                })
                .show();
    }



    /**
     * Requests a single permission.
     */
    private void showRequestPermissionDialog(Activity activity, String permission, int resultCode) {
        ActivityCompat.requestPermissions(activity, new String[]{permission}, resultCode);
    }

    /**
     * Requests multiple permissions in one call.
     */
    private void showRequestPermissionsDialog(Activity activity, String[] permissions,
                                              int resultCode) {
        ActivityCompat.requestPermissions(activity, permissions, resultCode);
    }

    /**
     * Returns a message to be shown to the user that explains why a specific permission is
     * required.
     */
    public String messageForRationale(String permission) {
        String s;
        switch (permission) {
            case Manifest.permission.READ_PHONE_STATE:
                s = "access this device's state";
                break;
            case Manifest.permission.ACCESS_FINE_LOCATION:
                s = "access the location of this device";
                break;
            case Manifest.permission.SEND_SMS:
                s = "send text messages";
                break;
            default:
                throw new IllegalArgumentException("Permission not handled: " + permission);
        }

        return String.format("MyApp needs permission to %s.", s);
    }

    /**
     * Get the RequestCode for the given permission.
     */
    public int requestCodeForPermission(String permission) {
        int code;
        switch (permission) {
            case Manifest.permission.READ_PHONE_STATE:
                code = RequestCode.PERMISSION_READ_PHONE_STATE;
                break;
            case Manifest.permission.ACCESS_FINE_LOCATION:
                code = RequestCode.PERMISSION_FINE_LOCATION;
                break;
            case Manifest.permission.SEND_SMS:
                code = RequestCode.PERMISSION_SEND_SMS;
                break;
            case Manifest.permission.WRITE_EXTERNAL_STORAGE:
                code = RequestCode.PERMISSION_WRITE_EXTERNAL_STORAGE;
                break;
            default:
                throw new IllegalArgumentException("Permission not handled: " + permission);
        }

        return code;
    }

    /**
     * Called if all requested permissions were granted in the same dialog.
     * E.g. FINE_LOCATION and SEND_SMS were requested, and both were granted.
     * <p>
     * Child class can override this method if it wants to know when this happens.
     * <p>
     * Linter can show an unjust "call requires permission" warning in child class if a method that
     * requires permission(s) is called. Silence it with `@SuppressWarnings("MissingPermission")`.
     */
    protected void onAllPermissionsGranted(String[] permissions) {

    }

    /**
     * Called for all permissions that were granted in the same dialog, in case not all were
     * granted. E.g. if FINE_LOCATION, COARSE_LOCATION and SEND_SMS were requested and FINE_LOCATION
     * was not granted but COARSE_LOCATION and SEND_SMS were, it will be called for COARSE_LOCATION
     * and SEND_SMS.
     * <p>
     * Child class can override this method if it wants to know when this happens.
     * <p>
     * Linter can show an unjust "call requires permission" warning in child class if a method that
     * requires permission(s) is called. Silence it with `@SuppressWarnings("MissingPermission")`.
     */
    protected void onPermissionGranted(String permission) {

    }

    /**
     * Called for all permissions that were denied in the same dialog, handled one by one.
     * <p>
     * Child class should not override this general behavior.
     */
    protected void onPermissionDenied(String permission) {
        String message = messageForRationale(permission);
        showRationaleDialog(this, permission, message, requestCodeForPermission(permission));
    }

    /**
     * Called for all permissions that were blocked in the same dialog, handled one by one.
     * <p>
     * Blocked means a user denied a permission with the 'never ask again' checkbox checked.
     * <p>
     * Child class must override and decide what to do when a permission is blocked.
     */
    protected abstract void onPermissionBlocked(String permission);
}
