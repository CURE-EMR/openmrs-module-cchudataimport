/**
 * This Source Code Form is subject to the terms of the Mozilla Public License,
 * v. 2.0. If a copy of the MPL was not distributed with this file, You can
 * obtain one at http://mozilla.org/MPL/2.0/. OpenMRS is also distributed under
 * the terms of the Healthcare Disclaimer located at http://openmrs.org/license.
 *
 * Copyright (C) OpenMRS Inc. OpenMRS is a registered trademark and the OpenMRS
 * graphic logo is a trademark of OpenMRS Inc.
 */
package org.openmrs.module.cchudataimport.web.controller;

import java.io.IOException;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.openmrs.User;
import org.openmrs.api.UserService;
import org.openmrs.module.cchudataimport.api.csv.reader.SimpleExcelReader;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * This class configured as controller using annotation and mapped with the URL of
 * 'module/${rootArtifactid}/${rootArtifactid}Link.form'.
 */
@Controller("${rootrootArtifactid}.CCHUDataImportController")
@RequestMapping(value = "module/${rootArtifactid}/${rootArtifactid}.form")
public class CCHUDataImportController {
	
	/** Logger for this class and subclasses */
	protected final Log log = LogFactory.getLog(getClass());
	
	@Autowired
	UserService userService;
	
	/** Success form view name */
	private final String VIEW = "/module/${rootArtifactid}/${rootArtifactid}";
	
	/**
	 * Initially called after the getUsers method to get the landing form name
	 * 
	 * @return String form view name
	 */
	@RequestMapping("/module/cchucsvutils/addObs")
	public String onGet() throws IOException {
		SimpleExcelReader r = new SimpleExcelReader();
		r.prepareObsImport();
		return VIEW;
	}
	
	@RequestMapping("/module/cchucsvutils/addDates")
	public String addDate() throws IOException {
		SimpleExcelReader r = new SimpleExcelReader();
		r.addDateofBirth();
		return VIEW;
	}
	
	@RequestMapping("/module/cchucsvutils/addIds")
	public String addIds() throws IOException {
		SimpleExcelReader r = new SimpleExcelReader();
		r.addIds();
		return VIEW;
	}
	
	@RequestMapping("/module/cchucsvutils/addRegistrationDiagnosis")
	public String addRegistrationDiagnosis() throws IOException {
		SimpleExcelReader r = new SimpleExcelReader();
		r.addRegistrationDiagnosis();
		return VIEW;
	}
	
	@RequestMapping("/module/cchucsvutils/addPersonAttributes")
	public String addPersonAttributes() throws IOException {
		log.error("=======================Ndi muri controller");
		SimpleExcelReader r = new SimpleExcelReader();
		r.preparePersonAttributesImport();
		return VIEW;
	}
	
	@RequestMapping("/module/cchucsvutils/saveAllPatients")
	public String saveAllPatients() throws IOException {
		SimpleExcelReader r = new SimpleExcelReader();
		r.saveAllPatients();
		return VIEW;
	}
	
	@RequestMapping("/module/cchucsvutils/saveAllNullDoBs")
	public String saveAllNullDoBs() throws IOException {
		SimpleExcelReader r = new SimpleExcelReader();
		r.saveAllNullDoBs();
		return VIEW;
	}
	
	@RequestMapping("/module/cchucsvutils/addSurgicalProcedureObsGroup")
	public String addSurgicalProcedureObsGroup() throws IOException {
		SimpleExcelReader r = new SimpleExcelReader();
		r.addSurgicalProcedureObsGroup();
		return VIEW;
	}
	
	/**
	 * This class returns the form backing object. This can be a string, a boolean, or a normal java
	 * pojo. The bean name defined in the ModelAttribute annotation and the type can be just defined
	 * by the return type of this method
	 */
	@ModelAttribute("users")
	protected List<User> getUsers() throws Exception {
		List<User> users = userService.getAllUsers();
		
		// this object will be made available to the jsp page under the variable name
		// that is defined in the @ModuleAttribute tag
		return users;
	}
	
}
