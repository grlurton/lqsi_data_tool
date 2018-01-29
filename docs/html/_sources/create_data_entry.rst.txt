Setting-up a Data entry file
===============================

Downloading and unpacking
-------------------------

1. Download the most recent version of the files should be downloaded on `<https://github.com/grlurton/xlqsi/blob/master/xlqsi.zip>`_. The zip file contains the two modules, as well as the Python libraries necessary to run the different contents.
2. Copy the zipfile in the directory you want to use for your project.
3. Unzip the zipfile.

.. warning:: For all the data entry setup, it is important keep the data entry file in the same directory as the *build* folder, so that it can access the necessary libraries.

Once  you open the Excel files, you will need to allow macros to be able to use some of the functions.

.. figure:: images/security_check.png
  :name: security_check

  Allowing macros (french translation)



Set language
-------------

It is currently possible to use XLQSI in five different languages. The English, French, Russian and Spanish versions have been directly taken from the WHO's website. The Khmer translation was made by ITech's team in Cambodia.

Language selection happens on the first tab of the data entry workbook, just by choosing the desired language in the dropdown menu as shown in :numref:`translation`.

.. figure:: images/translation_image.png
  :name: translation

  Setting the translation

Adding a new translation language to XLQSI is a simple even if very systematic task. The procedure to adapt translations or add new ones is described in :doc:`technical documentation<technical>`.


.. _define-the-completion-checks:

Define the completion checks
-----------------------------

To get more background on the difference between the *checklist items* and the *completion checks* please refer to the :doc:`introductory content<intro>`.

Completion checks are defined by the users to adapt generic checklist items to a specific context. For example, one of the checklist items defined by WHO are "Are sharps securely and safely incinerated without the risk for needle accidents?". In Cambodia, for this item to be validated, it necessitated the completion of three completion checks :
1. All sharps wastes are stored in sharps container,
2. Sharps containers are autoclaved and
3. Sharps containers are incinerated properly

We encourage each user to think on how much they need to adapt completion checks to their own context. For users using XLQSI in the context of the implementation of national policies, they should refer to national norms and procedures to accurately define completion checks. Additionally, different completion checks may have to be defined for different types of laboratories. For examples of how completion checks can be defined, we share the example of how ITech has been defining them in `Cambodia in XLSQI's repository examples <https://github.com/grlurton/xlqsi/tree/master/examples>`_.

.. note:: It is not necessary to define completion checks for all checklist items. For checklist items that are directly applicable to the user's situation, the user will be able to validate the item directly during data entry.

The definition of the completion checks is made in two steps:

1. In each *Criterias* sheet (*Phase1 - Criterias*, *Phase2 - Criterias*, *Phase3 - Criterias* and *Phase4 - Criterias*), in the *Completion Check* columns, the user can define the completion checks he needs. In :numref:`define_completion_check`, we can see how the checklist item *Have critical environmental/equipment parameters been identified?* was adapted into four completion checks.

.. figure:: images/define_completion_check.png
  :name: define_completion_check

  Examples of completion checks for a given checklist item

Also, you can access each of the sheet through the links in the dedicated box in the  *Getting Started* sheet (see :numref:`click_completion`).

2. The goal of XLQSI is to make data collection easier. Once all the completion checks have been defined for each phase, you can thus import them in synthetic lists for each phase. This is done by clicking on the *Extract Criterias Checklist* in the *Getting Started* sheet. On the right side of :numref:`click_completion`, you can see a counter. This count shows how many completion checks have been imported. Wait until the counter stops going up before modifying completion checks.

.. figure:: images/import_checklist.png
  :name: click_completion

  Completion checks importation

The completion of these lists is explained in the section on :doc:`data entry<fill_data_entry>`.

Freezing the data entry file
------------------------------

It may be useful to freeze the data entry file, to prevent undue modification by users, or to make its use easier. In the current stage of XLQSI, there is no systematic way to do it, and we welcome feedback of users regarding which aspect of the file could be deleted or frozen (this can be communicated through the `dedicated github issue <https://github.com/grlurton/xlqsi/issues/2>`_ ).
