#pragma once
#include "../API/iContext.cl.h"
#include "../ComLightLib/comLightServer.h"
#include "WhisperContext.h"
#include "Spectrogram.h"
#include "TranscribeResult.h"
#include "sTokenData.h"
#include "../ML/Device.h"
#include <random>

namespace Whisper
{
	class ContextImpl : public ComLight::ObjectRoot<iContext>
	{
		const DirectCompute::Device& device;
		const WhisperModel& model;
		ComLight::CComPtr<iModel> modelPtr;
		DirectCompute::WhisperContext context;
		Spectrogram spectrogram;
		int64_t mediaTimeOffset = 0;
		iSpectrogram* currentSpectrogram = nullptr;
		class CurrentSpectrogramRaii;
		ProfileCollection profiler;

		HRESULT COMLIGHTCALL getModel( iModel** pp ) override final;
		HRESULT COMLIGHTCALL timingsPrint() override final;
		HRESULT COMLIGHTCALL timingsReset() override final;
		HRESULT COMLIGHTCALL fullDefaultParams( eSamplingStrategy strategy, sFullParams* rdi ) override final;
		HRESULT COMLIGHTCALL runFullImpl( const sFullParams& params, const sProgressSink& progress, iSpectrogram& mel );
		HRESULT COMLIGHTCALL runFull( const sFullParams& params, const iAudioBuffer* buffer ) override final;
		HRESULT COMLIGHTCALL runStreamed( const sFullParams& params, const sProgressSink& progress, const iAudioReader* reader ) override final;
		HRESULT COMLIGHTCALL runCapture( const sFullParams& params, const sCaptureCallbacks& callbacks, const iAudioCapture* reader ) override final;

		struct Segment
		{
			int64_t t0;
			int64_t t1;
			mutable std::string text;
			std::vector<sTokenData> tokens;
			size_t memoryUsage() const;
		};
		std::vector<Segment> result_all;

		std::vector<whisper_token> prompt_past;

		// [EXPERIMENTAL] token-level timestamps data
		int64_t t_beg = 0;
		int64_t t_last = 0;
		whisper_token tid_last = 0;
		std::vector<float> energy; // PCM signal energy

		// [EXPERIMENTAL] speed-up techniques
		int32_t exp_n_audio_ctx = 0; // 0 - use default

		HRESULT encode( iSpectrogram& mel, int seek );
		HRESULT decode( const int* tokens, size_t length, int n_past, int threads );
		HRESULT detectLanguage( iSpectrogram& mel, int seek, int threads, uint32_t& language );
		sTokenData sampleBest( const float* probs, bool force_timestamp, bool is_initial, float temperature );
		sTokenData sampleBest( float temperature );
		sTokenData sampleTimestamp( bool initial, float temperature );

		// Random source for temperature sampling (temperature > 0). Seeded deterministically
		// so transcripts are reproducible across runs.
		std::mt19937 rng{ 0 };

		// Result of decoding a single segment at one temperature, plus the quality metrics
		// used to decide whether to fall back to a higher temperature.
		struct SegmentDecode
		{
			std::vector<sTokenData> tokens; // decoded tokens, already resized to result_len
			int seek_delta = 0;
			bool failed = false;            // decoder never found a usable end-of-segment
			double avg_logprob = 0.0;       // mean log-probability of the kept tokens
			double entropy = 0.0;           // token-id entropy over the last 32 tokens
			float no_speech_prob = 0.0f;    // probability of the no-speech token at the first step
		};
		HRESULT decodeSegment( const sFullParams& params, int seek, int seek_end,
			const std::vector<whisper_token>& prompt_init, float temperature, SegmentDecode& out );
		static double computeEntropy( const std::vector<sTokenData>& tokens );

		// Non-speech / blank token suppression, mirrors whisper.cpp whisper_process_logits
		std::vector<int> suppress_tokens; // non-speech token ids, built once from the vocabulary
		int token_space = -1;             // id of the " " (blank) token, or -1 if absent
		bool suppress_built = false;
		bool suppress_blank = false;
		bool suppress_nst = false;
		void buildSuppressTokens();
		
		int wrapSegment( int max_len );
		void expComputeTokenLevelTimestamps( int i_segment, float thold_pt, float thold_ptsum );

		std::vector<float> probs;
		std::vector<std::pair<double, Vocabulary::id>> probs_id;

		mutable TranscribeResultStatic results;

		HRESULT COMLIGHTCALL makeResults( eResultFlags flags, TranscribeResult& res, bool moveStrings ) const noexcept;

		HRESULT COMLIGHTCALL getResults( eResultFlags flags, iTranscribeResult** pp ) const noexcept override final;
		HRESULT COMLIGHTCALL detectSpeaker( const sTimeInterval& time, eSpeakerChannel& result ) const noexcept override final;

		int defaultThreadsCount() const;

		__m128i getMemoryUse() const;
		mutable std::vector<StereoSample> diarizeBuffer;

	public:

		ContextImpl( const DirectCompute::Device& dev, const WhisperModel& modelData, iModel* modelPointer );
	};
}