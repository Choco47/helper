//1.

const list = get(data['applicant'], 'fieldlist', []);

for (let i = 0; i < list.length; i++) {
	mapping[`pdf_hkid${i}`] = `${list[i].hkid.id}(${list[i].hkid.checkDigit})`;
}



//2.

const list = get(data['applicant'], 'qc_reg', []);

const mapping = {
      ...flattenToStringMap(data),
      'applicant.qc_reg.0.hkid': isEmpty(list[0]) ? '' : `${list[0].hkid.id}(${list[0].hkid.checkDigit})`,
      'applicant.qc_reg.1.hkid': isEmpty(list[1]) ? '' : `${list[1].hkid.id}(${list[1].hkid.checkDigit})`,
      'applicant.qc_reg.2.hkid': isEmpty(list[2]) ? '' : `${list[2].hkid.id}(${list[2].hkid.checkDigit})`,
