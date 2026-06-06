#pragma once
#include "audioConstants.h"

namespace Whisper
{
	struct alignas( 8 ) StereoSample
	{
		float left, right;
	};

	__interface iSpectrogram
	{
       // Make a buffer with length * getMelCount() floats, starting at the specified offset
		// An implementation of this interface may visualize the spectrogram, making pieces on demand
		HRESULT makeBuffer( size_t offset, size_t length, const float** buffer, size_t& stride );

		size_t getMelCount() const;

		// Apparently, the length unit is 160 input samples = 10 milliseconds of audio
		size_t getLength() const;

		// If the source data is stereo, copy the specified slice of the data into the provided vector
		HRESULT copyStereoPcm( size_t offset, size_t length, std::vector<StereoSample>& buffer ) const;
	};

	// RAII class to deal with iSpectrogram's makeBuffer method.
	// Throws exceptions when things fail.
	class MelBufferRaii
	{
		const float* pointer;
		size_t stride;
     size_t melCount;
	public:

		HRESULT make( iSpectrogram& mel, size_t off, size_t len )
		{
            melCount = mel.getMelCount();
			return mel.makeBuffer( off, len, &pointer, stride );
		}

		const float* operator[]( size_t idx ) const
		{
          assert( idx < melCount );
			return pointer + idx * stride;
		}

		const BYTE* bytePtr() const { return (const BYTE*)pointer; }
		LONG strideBytes() const { return (LONG)stride * 4; }
	};
}