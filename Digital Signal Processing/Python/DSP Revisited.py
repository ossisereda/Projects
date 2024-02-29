"""
Digital Signal Processing
Creator: Ossi Sereda
Student ID: 6614050044

Merged and generalized code from a personal project work
(Audio Signal Denoising for Speech Intelligibility) for the course
Digital Signal Processing (Kasetsart University)

The code works optimally if the used audio files are 32 bit wav.-files
with 44,1 kHz sample rate

Some of the functions are for demonstration purposes only (for example, the
destructive interference only works for an optimized set of files), so it is
recommended to check the provided project powerpoint and README files
to get a full idea of what the functions are made for.
"""

from scipy.io import wavfile
from scipy.stats import pearsonr
import matplotlib.pyplot as plt
from scipy.signal import butter, lfilter
import numpy as np


def get_clean(speech_file, noise_file, output_file, gain):
    """
    Cleans a noisy speech signal with a known noise profile. Cleaning is done
    via inverted polarity of the noise spectrum.

    :param speech_file: Directory for an audio file with speech and a known
                        noise profile
    :param noise_file: Directory for an audio file with the known noise profile
    :param output_file: Directory for the cleaned speech signal with no noise
    :param gain: Applies gain based on a given number. Essentially a volume
                 knob for the output file
    """
    # Read (noisy) speech and noise files
    _, noisy_speech_frames = wavfile.read(speech_file)
    _, noise_frames = wavfile.read(noise_file)

    # Add phase inverted noise to the noisy speech and apply additional gain
    combined_frames = noisy_speech_frames + (-noise_frames)
    combined_frames *= gain

    # Save the result to a new WAV file
    wavfile.write(output_file, rate=44100, data=combined_frames)


def calculate_correlation(first, second):
    """
    Calculates the correlation between two different audio files,
    helpful for example when determining the effect of destructive interference
    :param first: Directory for the first file to be compared
    :param second: Directory for the second file to be compared
    :return: Pearsons correlation coefficient, number between -1 and 1
    """

    # Read speech and noise files
    _, speech_frames = wavfile.read(first)
    _, noise_frames = wavfile.read(second)

    # Ensure both signals have the same length
    min_length = min(len(speech_frames), len(noise_frames))
    speech_frames = speech_frames[:min_length]
    noise_frames = noise_frames[:min_length]

    # Calculate Pearson correlation coefficient
    correlation_coefficient, _ = pearsonr(speech_frames, noise_frames)

    return correlation_coefficient


def plot_spectrum(input_file, title, sample_rate=44100):
    """
    Plots the frequency-magnitude spectrum of a given wav.-file
    in the human hearing range (20 Hz - 20 kHz)

    :param input_file: Frames of the wav.-file to be plotted
    :param title: Title for the graph
    :param sample_rate: Sample rate of the signal, suggested to be kept at 44,1 kHz
    """

    _, data = wavfile.read(input_file)

    # Compute the Fast Fourier Transform (FFT), and
    # the frequencies corresponding to the FFT values
    spectrum = np.fft.fft(data)
    frequencies = np.fft.fftfreq(len(spectrum), d=1/sample_rate)

    # Plot the frequency-magnitude -spectrum
    plt.figure()
    plt.title(title)
    plt.xlabel('Frequency (Hz)')
    plt.ylabel('Magnitude')
    plt.plot(frequencies, np.abs(spectrum))

    # Set the frequency axis to logarithmic scale for better readability,
    # and limit the range to 20 Hz - 20 kHz
    plt.xscale('log')
    plt.xlim(20, 20000)

    plt.show()


def butter_highpass(data, cutoff_freq, sample_rate, order=4):
    """
    Creates a Butterworth high-pass filter based on theory and characteristics
    of the specified input file

    :param data: frames of the input audio file
    :param cutoff_freq: arbitrary cutoff frequency for the high-pass filter
    :param sample_rate: sample rate of the input audio file
    :param order: arbitrary order for the filter, affects the roll-off
                  characteristics of the filter
    """

    nyquist = 0.5 * sample_rate
    normal_cutoff = cutoff_freq / nyquist
    b, a = butter(order, normal_cutoff, btype='high', analog=False)[:2]
    return lfilter(b, a, data)


def apply_butter(input_file, output_file, cutoff_freq):
    """
    Reads the file to be cleaned, applies a high-pass filter to the file
    and saves the result to a new wav.-file

    :param input_file: Directory to the wav.-file, to which the filter is applied
    :param output_file: Directory to the output file
                        (= input file with the filter applied)
    :param cutoff_freq: Arbitrary cutoff frequency for the high-pass filter
    """

    sample_rate, input_frames = wavfile.read(input_file)
    output_frames = butter_highpass(input_frames, cutoff_freq, sample_rate)
    wavfile.write(output_file, sample_rate, output_frames.astype(np.float32))


def continue_or_quit():
    """
    Creates a prompt for the user to answer, if they want to continue using
    the operations provided, or quit the program

    :return: boolean value, which will either break or redo the loop in
             main function
    """

    while True:
        yesno = input("Do you want to do another operation? (Y/N) ").upper()
        if yesno == "Y":
            return True
        elif yesno == "N":
            return False
        else:
            print("Invalid input, write 'Y' or 'N'.")


def main():

    print("\nChoose operation:\n")
    print("Destructive interference (type DEST)")
    print("Correlation between two signals (type CORR)")
    print("Frequency-magnitude spectrum of a signal (type PLOT)")
    print("High-pass filter (type FILT)")

    # Loop to keep the program going, until the user decides to do no more
    # operations
    while True:
        user_input = input("\nOperation: ").upper()

        # Destructive interference (use idealized files provided)
        if user_input == "DEST":
            print("\nDestructive interference:")
            file1 = input("Type the directory to the first file: ")
            file2 = input("Type the directory to the second file: ")
            output = input("Type the directory to where you want the cleaned "
                           "signal to be sent to: ")
            while True:
                gain_input = input(
                    "How loud do you want the output signal to be? (enter a number): ")
                try:
                    gain = float(gain_input) if gain_input else 1.0
                    break
                except ValueError:
                    print("Invalid input. Please enter a valid number.")

            try:
                get_clean(file1, file2, output, gain)
                print("\nSignal cleaned succesfully\n")
                pass
            except FileNotFoundError:
                print("\nThe specified file(s) was not found. Please check "
                      "the file path(s).\n")

            if not continue_or_quit():
                break

        # Pearson correlation coefficient between two signals
        elif user_input == "CORR":
            print("\nCorrelation between two signals:")
            file1 = input("Type the directory to the first file: ")
            file2 = input("Type the directory to the second file: ")

            try:
                print(f"\nPearsons correlation coefficient for the two files "
                      f"is: {calculate_correlation(file1, file2):.4f}\n")
                pass
            except FileNotFoundError:
                print("\nThe specified file(s) was not found. Please check "
                      "the file path(s).\n")

            if not continue_or_quit():
                break

        # Plot the frequency-magnitude spectrum for a desired wav.-file
        elif user_input == "PLOT":
            print("\nFrequency-magnitude spectrum:")
            file = input("Type the directory to the signal for which you want "
                         "the spectrum: ")
            name = input("Name the plot for the spectrum: ")

            try:
                plot_spectrum(file, name)
                print("\nSpectrum printed succesfully\n")
                pass
            except FileNotFoundError:
                print("\nThe specified file(s) was not found. Please check "
                      "the file path(s).\n")

            if not continue_or_quit():
                break

        # Apply Butterworth filter
        elif user_input == "FILT":
            print("\nHigh-pass filter:")
            input_directory = input("Type the directory for the signal for "
                                    "which you want to apply the filter: ")
            output_directory = input("Type the directory for the output signal "
                                     "(input signal with the filter applied): ")
            cutoff_freq = int(input("Type the desired cutoff frequency (Hz): "))

            try:
                apply_butter(input_directory, output_directory, cutoff_freq)
                print("\nHigh-pass filter applied succesfully\n")
                pass
            except FileNotFoundError:
                print("\nThe specified file(s) was not found. Please check "
                      "the file path(s).\n")
            except ValueError:
                print("\nThe specified cutoff frequency was not of accepted "
                      "value.\n")

            if not continue_or_quit():
                break

        else:
            print("\nInvalid operation (type DEST / CORR / PLOT / FILT)")

    # Quitting the program
    print("\nHave a nice day!")


if __name__ == "__main__":
    main()
