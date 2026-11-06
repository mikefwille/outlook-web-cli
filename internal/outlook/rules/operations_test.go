package rules

import "testing"

func TestRuleFingerprintDetectsChanges(t *testing.T) {
	a := []byte(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"display","RawIdentity":"id"},"Name":"Example","Enabled":false}`)
	b := []byte(`{"Name":"Example","Enabled":false,"Identity":{"RawIdentity":"id","DisplayName":"display"},"__type":"InboxRule:#Exchange"}`)
	c := []byte(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"display","RawIdentity":"id"},"Name":"Changed","Enabled":false}`)
	x, err := Parse(a)
	if err != nil {
		t.Fatal(err)
	}
	y, err := Parse(b)
	if err != nil || x.Version != y.Version {
		t.Fatal("fingerprint depends on key order")
	}
	z, err := Parse(c)
	if err != nil || x.Version == z.Version {
		t.Fatal("fingerprint missed a change")
	}
	if x.SafeToRename() {
		t.Fatal("incomplete rule was editable")
	}
}
