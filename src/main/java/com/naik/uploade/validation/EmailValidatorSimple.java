package com.naik.uploade.validation;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

// Email validation method
    public class EmailValidatorSimple {

        private static final String EMAIL_PATTERN = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,6}$";

        private static final Pattern pattern = Pattern.compile(EMAIL_PATTERN, Pattern.CASE_INSENSITIVE);

        public static boolean isValid(final String email) {
            Matcher matcher = pattern.matcher(email);
            return matcher.matches();
        }

    }

