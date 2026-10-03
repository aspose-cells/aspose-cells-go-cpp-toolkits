package tests

import (
	"testing"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/formats"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/saveoptions"
)

// stubOption is a minimal SaveOption used to exercise the formats registry
// without depending on any saveoptions implementation.
type stubOption struct{}

func (stubOption) Apply(source []byte) ([]byte, error) { return source, nil }
func (stubOption) GetFormat() string                   { return "stub" }

// The fake extensions below must not collide with real registered formats,
// so a distinctive "zzz-" prefix is used.
func TestRegistryCaseInsensitive(t *testing.T) {
	formats.Register("zzz-test-ext", func() saveoptions.SaveOption { return stubOption{} })
	defer formats.Unregister("zzz-test-ext")

	for _, key := range []string{"zzz-test-ext", "ZZZ-TEST-EXT", "ZzZ-tEsT-eXt"} {
		opt := formats.Get(key)
		if opt == nil {
			t.Fatalf("Get(%q) returned nil", key)
		}
		if got := opt.GetFormat(); got != "stub" {
			t.Errorf("Get(%q).GetFormat() = %q, want %q", key, got, "stub")
		}
	}
}

func TestRegistryGetUnknownReturnsNil(t *testing.T) {
	if opt := formats.Get("definitely-not-a-format"); opt != nil {
		t.Errorf("Get(unknown) = %v, want nil", opt)
	}
}

func TestListContainsRegistered(t *testing.T) {
	formats.Register("zzz-list-check", func() saveoptions.SaveOption { return stubOption{} })
	defer formats.Unregister("zzz-list-check")
	found := false
	for _, ext := range formats.List() {
		if ext == "zzz-list-check" {
			found = true
			break
		}
	}
	if !found {
		t.Errorf("List() = %v, want it to contain zzz-list-check", formats.List())
	}
}

// TestUnregisterRemovesExtension verifies Unregister cleans the registry, which
// is how tests avoid polluting the process-wide registry with fake extensions.
func TestUnregisterRemovesExtension(t *testing.T) {
	formats.Register("zzz-unreg", func() saveoptions.SaveOption { return stubOption{} })
	if formats.Get("zzz-unreg") == nil {
		t.Fatal("zzz-unreg should be registered before Unregister")
	}
	formats.Unregister("zzz-unreg")
	if formats.Get("zzz-unreg") != nil {
		t.Error("zzz-unreg should return nil after Unregister")
	}
}

func TestListSorted(t *testing.T) {
	exts := formats.List()
	for i := 1; i < len(exts); i++ {
		if exts[i-1] > exts[i] {
			t.Fatalf("List() is not sorted: %q > %q at index %d", exts[i-1], exts[i], i)
		}
	}
}
