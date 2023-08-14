# GaussSum (http://gausssum.sf.net)
# Copyright (C) 2006-2013 Noel O'Boyle <baoilleach@gmail.com>
#
# This program is free software; you can redistribute and/or modify it
# under the terms of the GNU General Public License as published by the
# Free Software Foundation; either version 2, or (at your option) any later
# version.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY, without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# General Public License for more details.

import math
import numpy
from cclib.parser import ccopen


def broadenSpectrum(start, end, numpts, peaks, width, formula):
    """
    Broadens spectrum data. Creates a distribution function around
    each peak (pos, height) and adds up the contributions from each 
    distribution over numpts from start to end to create a spectrum
    that looks closer to one obtained by experiment

    
    formula is a function such as gaussianpeak or delta
    """
    spectrum = numpy.zeros(numpts,"d")
    xvalues = numpy.linspace(start, end, numpts)
    print(xvalues, start)
    for i in range(numpts):
        x = xvalues[i]
        for pos, height in peaks:
            spectrum[i] = spectrum[i] + formula(x, pos, height, width)

    return xvalues, spectrum



def lorentzian(x, peak, height, width):
    """The lorentzian curve.

    f(x) = a/(1+a)

    where a is FWHM**2/4
    """
    a = width**2./4.
    return float(height)*a/( (peak-x)**2 + a )

    
def activity_to_intensity(activity, frequency, excitation, temperature):
    """Convert Raman acitivity to Raman intensity according to
    Krishnakumar et al, J. Mol. Struct., 2004, 702, 9."""

    excitecm = 1 / (1e-7 * excitation)
    f = 1e-13
    above = f * (excitecm - frequency)**4 * activity
    exponential = -6.626068e-34 * 299792458 * frequency / (1.3806503e-23 * temperature)
    below = frequency * (1 - math.exp(exponential))
    return above / below


def parseFile(inputFileName):
    mode = []
    freq = []
    act = []
    with open(inputFileName, "r") as f:
        for line in f.readlines()[1:]:
            numbers = line.split("\t")
            if len(numbers) == 3:
                mode.append(int(numbers[0]))
                freq.append(float(numbers[1]))
                act.append(float(numbers[2]))

    return mode, freq, act

def ramanSpectra(inputFileName, outputFileName, start, end, numpts, FWHM, scaleFunction, excitation, temperature):
    print("Parsing file")

    mode, freq, act = parseFile(inputFileName)
    unscaledFreq = freq.copy()
    scale = []
    
    for i in range(len(freq)):
        scalingFactor = scaleFunction(freq[i])
        freq[i] *= scalingFactor
        scale.append(scalingFactor)
        
    intensity = [activity_to_intensity(a, f, excitation, temperature) for f, a in zip(freq, act)]
    
    print("Broadening spectrum")
    xvalues, activity_spectrum = broadenSpectrum(start, end, numpts, list(zip(freq, act)), FWHM, lorentzian)
    xvalues, intensity_spectrum = broadenSpectrum(start, end, numpts, list(zip(freq, intensity)), FWHM, lorentzian)
        
    
    
    print("Writing scaled spectrum to", outputFileName) 
    with open(outputFileName, "w") as outputFile:
        outputFile.write("\t".join(["Spectrum Freq", "Spectrum Activity", "Spectrum Intensity", 
            "Mode", "Unscaled Freq", "Scale", "Scaled Freq", "Activity", "Intensity"]))
        outputFile.write("\n")
        
        i = 0
        while i < max(numpts, len(freq)):
            if i < numpts:  # write the spectrum data
                if activity_spectrum[i] < 1e-20:
                    activity_spectrum[i] = 0.
                if intensity_spectrum[i] < 1e-20:
                    intensity_spectrum[i] = 0.
                print(intensity_spectrum[i])
                outputFile.write(str(xvalues[i]) + "\t" + str(activity_spectrum[i]) + "\t" + str(intensity_spectrum[i]))
            else:
                outputFile.write("\t\t")
                
            
            if i < len(freq):
                print(mode[i])
                outputFile.write("\t" + str(mode[i]) + "\t" + str(unscaledFreq[i]) + "\t" + str(scale[i]) + "\t"
                    + str(freq[i]) + "\t" + str(act[i]) + "\t" + str(intensity[i]))

            outputFile.write("\n")
            i += 1
    print(intensity)
    return xvalues, activity_spectrum, intensity_spectrum


if __name__ == "__main__":
    inputFileName = "vasp_raman.dat"
    outputFileName = "test.txt"
    start = 80 
    end = 900
    numpts = 4000
    FWHM = 10
    excitation = 10
    temperature = 273
    
    def scaleFunction(freq):
        if freq < 1111.11:
            scalingFactor = 0.979
        elif freq > 2500:
            scalingFactor = 0.961
        else:
            scalingFactor = 0.973    
        return scalingFactor
        
    
    ramanSpectra(inputFileName, outputFileName, start, end, numpts, FWHM, scaleFunction, excitation, temperature)
