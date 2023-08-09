def reverse(x):
    """
    This function will return the reverse of a
    string input.
    """
    return x[::-1]
    """
    The time complextity of this
    is O(n) because we are using string splicing
    to access the equivalent of a C pointer to the
    last item in our string and then decrementing
    that pointer until the last item is accessed.
    This means that there is also no new allocation
    involved.
    """

print(reverse("ABC"))
print(reverse("123"))
print(reverse("Harrison Weiss"))