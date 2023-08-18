package com.jerry.numbers.util;

public class LongRandomNumberGenerator {
    public static void main(String[] args) {

//        working code to generate long random numbers with lengh between 14 and 16
        for (int i = 1; i <= 20; i++) {
            long randomNumber = (long) (Math.random() * Math.pow(10, (int) (Math.random() * 6) + 14));
            int sizeOfRandomNumber = (int)Math.log10(randomNumber)+1;
            System.out.println(sizeOfRandomNumber +"\t"+randomNumber);
        }
    }
}
