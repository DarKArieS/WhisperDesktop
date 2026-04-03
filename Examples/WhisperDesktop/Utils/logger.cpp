#include "stdafx.h"
#include "logger.h"
#include "miscUtils.h"

using namespace Whisper;

void printTime( CStringA& rdi, Whisper::sTimeSpan time, bool comma, bool printFullSeconds )
{
	Whisper::sTimeSpanFields fields = time;
	const uint32_t hours = fields.days * 24 + fields.hours;
	const char separator = comma ? ',' : '.';
	if (printFullSeconds) {
		rdi.AppendFormat("%02d:%02d:%02d%c%03d (%02ds)",
			(int)hours,
			(int)fields.minutes,
			(int)fields.seconds,
			separator,
			fields.ticks / 10'000,
			(int)fields.fullSeconds);
	}
	else {
		rdi.AppendFormat("%02d:%02d:%02d%c%03d",
			(int)hours,
			(int)fields.minutes,
			(int)fields.seconds,
			separator,
			fields.ticks / 10'000 );
	}
}

HRESULT logNewSegments( const iTranscribeResult* results, size_t newSegments, bool printSpecial )
{
	sTranscribeLength length;
	CHECK( results->getSize( length ) );

	const size_t len = length.countSegments;
	size_t i = len - newSegments;

	const sSegment* const segments = results->getSegments();
	const sToken* const tokens = results->getTokens();

	CStringA str;
	for( ; i < len; i++ )
	{
		const sSegment& seg = segments[ i ];
		str = "[ ";
		printTime( str, seg.time.begin, false, true );
		str += " --> ";
		printTime( str, seg.time.end, false, true );

		// Compute average token probability and average timestamp token probability
		if (tokens != nullptr && seg.countTokens > 0)
		{
			float sumProb = 0.0f;
			float sumProbTs = 0.0f;
			uint32_t count = 0;
			for (uint32_t t = seg.firstToken; t < seg.firstToken + seg.countTokens; t++)
			{	
				const sToken& tok = tokens[t];

			
				if (!printSpecial && (tok.flags & eTokenFlags::Special))
					continue;
				sumProb += tok.probability;
				// sumProbTs += tok.probabilityTimestamp;
				// logInfo(u8"%.3f,%.3f,%.3f,%.3f", tok.probability, tok.probabilityTimestamp, tok.ptsum, tok.vlen);

				count++;
			}
			if (count > 0)
			{
				str.AppendFormat("p=%.3f", sumProb / count);
			}
		}

		str += " ]  ";

		str += seg.text;

		logInfo( u8"%s", cstr( str ) );
	}

	return S_OK;
}