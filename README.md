# Excelerator
A Microsoft Excel calculation speed-up add in.

## What is it?

This is an Excel plug-in I developed in my spare time (circa 2014-2015) which increases the computation time for Excel workbooks by a factor of at least 12x (in most cases, a lot higher).

It does this through several steps:

* Rewriting the formulae into a more efficient computation tree to figure out higher-order parallelism
* Leveraging CUDA-enabled GPUs to more efficiently compute the now-better-parallelised formulae
* Efficient use of multi-threading to compute non-parallelised computation chains

## File/folder structure

The file and folder structure of this repo is, shamefully, a big mess. This is an old project of mine, cobbled up in my spare time that I haven't had the chance to clean up or develop further in any recent years.

## Contributions

I would love for anyone to become an active contributor to this project. While it does require an initial refactor and clean-up, the benefits it provides to Excel users are well worth it.

## Why am I doing this?

It's a project that I would love to see become active again. There's an interesting, and very practical, story as to how I came up with the idea for it and hate seeing it gather cobwebs.

If you would like to contribute, feel free to get in touch with me at boyanvs@gmail.com.
